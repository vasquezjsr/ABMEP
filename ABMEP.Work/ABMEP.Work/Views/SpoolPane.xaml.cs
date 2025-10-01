// SpoolPane.xaml.cs — FULL DROP-IN
// Creates ONLY a true MEP Fabrication schedule (Pipework/Ductwork), filtered to the assembly,
// formatted by the selected Schedule View Template, and placed 1/16" left & 1/8" top.
// No Part List is created.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using ABMEP.Work.Schedules;

namespace ABMEP.Work.Views
{
    public partial class SpoolPane : UserControl
    {
        private readonly UIApplication _uiApp;
        private readonly UIDocument _uiDoc;

        private readonly SpoolRequestHandler _handler;
        private readonly ExternalEvent _extEvent;

        public ObservableCollection<AssemblyRow> Assemblies { get; } = new ObservableCollection<AssemblyRow>();
        private SpoolSettings _settings = new SpoolSettings();

        private List<string> _titleBlocks = new List<string>();
        private List<string> _scheduleTemplates = new List<string>();
        private List<string> _tagTypes = new List<string>();
        private List<string> _viewportTypes = new List<string>();

        public SpoolPane(UIApplication uiApp)
        {
            InitializeComponent();

            _uiApp = uiApp;
            _uiDoc = uiApp != null ? uiApp.ActiveUIDocument : null;

            _handler = new SpoolRequestHandler();
            _extEvent = ExternalEvent.Create(_handler);

            Loaded += (s, e) =>
            {
                try
                {
                    TryLoadLogo();
                    LoadChoicesFromDoc();
                    LoadSettingsFromDoc();
                    LoadAssembliesFromActiveView();
                    gridAssemblies.ItemsSource = Assemblies;

                    HookRowChangeNotifications();
                    UpdateActionButtons();
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Assembly Manager - ABMEP Assembly Manager",
                        "Initialization error:\n\n" + ex.Message);
                }
            };

            btnRefreshAssemblies.Click += (s, e) =>
            {
                LoadAssembliesFromActiveView();
                UpdateActionButtons();
            };

            btnClearSel.Click += (s, e) =>
            {
                foreach (var r in Assemblies) r.IsSelected = false;
                if (gridAssemblies != null) gridAssemblies.SelectedItems.Clear();
                gridAssemblies.Items.Refresh();
                UpdateActionButtons();
            };

            btnCreate.Click += (s, e) => Run("create");
            btnRefresh.Click += (s, e) => Run("refresh");

            gridAssemblies.SelectionChanged += gridAssemblies_SelectionChanged;
        }

        private void btnOpenSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var dlg = new AssemblySettingsWindow(
                    _settings,
                    _titleBlocks,
                    _scheduleTemplates,
                    _tagTypes,
                    _viewportTypes,
                    GetLogoPath());

                try
                {
                    IntPtr hwnd = _uiApp != null ? _uiApp.MainWindowHandle : IntPtr.Zero;
                    new WindowInteropHelper(dlg) { Owner = hwnd };
                }
                catch { }

                bool? ok = dlg.ShowDialog();
                if (ok == true && dlg.Result != null)
                {
                    _settings = dlg.Result;

                    var saveOpts = new SpoolOptions
                    {
                        Action = "save-settings",
                        Settings = _settings
                    };
                    _handler.Set(_uiApp, saveOpts, null);
                    _extEvent.Raise();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Assembly Manager - ABMEP Assembly Manager",
                    "Settings window error:\n\n" + ex);
            }
        }

        private void HeaderSelectAll_Click(object sender, RoutedEventArgs e)
        {
            bool on = (sender as CheckBox) != null && ((CheckBox)sender).IsChecked == true;
            foreach (var row in Assemblies) row.IsSelected = on;
            gridAssemblies.Items.Refresh();
            UpdateActionButtons();
        }

        // ----------------- UI helpers -----------------

        private void UpdateActionButtons()
        {
            bool anyChecked = Assemblies.Any(a => a.IsSelected);
            bool anySelected = gridAssemblies != null && gridAssemblies.SelectedItems != null && gridAssemblies.SelectedItems.Count > 0;
            bool any = anyChecked || anySelected;

            if (btnCreate != null) btnCreate.IsEnabled = any;
            if (btnRefresh != null) btnRefresh.IsEnabled = any;
        }

        private void HookRowChangeNotifications()
        {
            Assemblies.CollectionChanged += (s, e) =>
            {
                if (e.NewItems != null)
                {
                    foreach (AssemblyRow r in e.NewItems)
                        r.PropertyChanged += delegate { UpdateActionButtons(); };
                }
                UpdateActionButtons();
            };
        }

        // ----------------- Loaders -----------------

        private void TryLoadLogo()
        {
            try
            {
                string path = GetLogoPath();
                if (File.Exists(path))
                {
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.UriSource = new Uri(path, UriKind.Absolute);
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    imgLogo.Source = bmp;
                }
            }
            catch { }
        }

        private string GetLogoPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                @"Autodesk\Revit\Addins\2024",
                "AsBuilt MEP Logo.png");
        }

        private void LoadAssembliesFromActiveView()
        {
            Assemblies.Clear();
            Document doc = _uiDoc != null ? _uiDoc.Document : null;
            if (doc == null || _uiDoc.ActiveView == null) return;

            var col = new FilteredElementCollector(doc, _uiDoc.ActiveView.Id)
                .OfClass(typeof(AssemblyInstance))
                .Cast<AssemblyInstance>();

            foreach (var ai in col)
            {
                var et = doc.GetElement(ai.GetTypeId()) as ElementType;
                string name = et != null ? et.Name : (ai.Name ?? ("Assembly " + ai.Id.Value));
                Assemblies.Add(new AssemblyRow
                {
                    IsSelected = false,
                    Name = name,
                    ElementId = ai.Id.Value,
                    MemberCount = ai.GetMemberIds().Count
                });
            }
        }

        private List<long> GetTargetAssemblyIds()
        {
            var ids = new HashSet<long>(Assemblies.Where(a => a.IsSelected).Select(a => a.ElementId));
            if (gridAssemblies != null && gridAssemblies.SelectedItems != null)
            {
                foreach (var obj in gridAssemblies.SelectedItems)
                {
                    var r = obj as AssemblyRow;
                    if (r != null) ids.Add(r.ElementId);
                }
            }
            return ids.ToList();
        }

        private void Run(string action)
        {
            var ids = GetTargetAssemblyIds();
            if (ids.Count == 0)
            {
                TaskDialog.Show("ABMEP Assembly Manager",
                    "No assemblies are checked or selected.\nCheck or select one or more assemblies, then try again.");
                UpdateActionButtons();
                return;
            }

            var opts = new SpoolOptions
            {
                Action = action,
                AssemblyElementIds = ids,
                Settings = _settings
            };
            _handler.Set(_uiApp, opts, null);
            _extEvent.Raise();
        }

        private void LoadChoicesFromDoc()
        {
            Document doc = _uiDoc != null ? _uiDoc.Document : null;
            if (doc == null) return;

            _titleBlocks = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .Select(fs => fs.Name).Distinct().OrderBy(n => n).ToList();

            _scheduleTemplates = new FilteredElementCollector(doc)
                .OfClass(typeof(View)).Cast<View>()
                .Where(v => v.IsTemplate && v.ViewType == ViewType.Schedule)
                .Select(v => v.Name).Distinct().OrderBy(n => n).ToList();

            _tagTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .Cast<ElementType>()
                .Where(NameLooksLikeTag)
                .Select(et => et.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n)
                .ToList();

            if (_tagTypes.Count == 0)
            {
                var names = new List<string>();
                var cats = new[]
                {
                    BuiltInCategory.OST_FabricationDuctworkTags,
                    BuiltInCategory.OST_FabricationPipeworkTags,
                    BuiltInCategory.OST_DuctTags,
                    BuiltInCategory.OST_PipeTags
                };
                foreach (var bic in cats)
                {
                    names.AddRange(new FilteredElementCollector(doc)
                        .OfCategory(bic)
                        .WhereElementIsElementType()
                        .Cast<ElementType>()
                        .Select(tp => tp.Name));
                }
                _tagTypes = names.Distinct(StringComparer.OrdinalIgnoreCase)
                                 .OrderBy(n => n)
                                 .ToList();
            }

            _viewportTypes = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Viewports)
                .WhereElementIsElementType()
                .Cast<ElementType>()
                .Select(tpe => tpe.Name)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n)
                .ToList();

            if (_viewportTypes.All(n => !n.Equals("No Title", StringComparison.OrdinalIgnoreCase)))
                _viewportTypes.Insert(0, "No Title");

            if (string.IsNullOrEmpty(_settings.TitleBlockName) && _titleBlocks.Count > 0)
                _settings.TitleBlockName = _titleBlocks[0];
            if (string.IsNullOrEmpty(_settings.ScheduleTemplateName) && _scheduleTemplates.Count > 0)
                _settings.ScheduleTemplateName = _scheduleTemplates[0];
            if (string.IsNullOrEmpty(_settings.TagTypeName) && _tagTypes.Count > 0)
                _settings.TagTypeName = _tagTypes[0];

            if (string.IsNullOrEmpty(_settings.ViewportTypeName))
            {
                var noTitle = _viewportTypes.FirstOrDefault(n => n.Equals("No Title", StringComparison.OrdinalIgnoreCase));
                _settings.ViewportTypeName = noTitle ?? (_viewportTypes.Count > 0 ? _viewportTypes[0] : "No Title");
            }
        }

        private void LoadSettingsFromDoc()
        {
            Document doc = _uiDoc != null ? _uiDoc.Document : null;
            if (doc == null) return;

            Schema schema = SpoolSettingsStorage.GetOrCreateSchema();
            Entity ent = doc.ProjectInformation.GetEntity(schema);
            if (ent.IsValid())
                _settings = SpoolSettingsStorage.Read(ent);
        }

        private static bool NameLooksLikeTag(ElementType et)
        {
            if (et == null) return false;
            if (et.Category == null) return false;
            if (et.Category.CategoryType != CategoryType.Annotation) return false;

            return NameHasTag(et.Category.Name)
                || NameHasTag(et.FamilyName)
                || NameHasTag(et.Name);
        }

        private static bool NameHasTag(string s)
        {
            return !string.IsNullOrEmpty(s)
                && s.IndexOf("tag", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void gridAssemblies_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateActionButtons();
        }
    }

    // -------------- Models & storage (unchanged) --------------

    public class AssemblyRow : INotifyPropertyChanged
    {
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected == value) return; _isSelected = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected))); }
        }
        public string Name { get; set; }
        public long ElementId { get; set; }
        public int MemberCount { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
    }

    public class SpoolSettings
    {
        public string TitleBlockName { get; set; } = "";
        public string ScheduleTemplateName { get; set; } = "";
        public string TagTypeName { get; set; } = "";
        public string ViewportTypeName { get; set; } = "";

        public bool View3D { get; set; } = true;
        public bool ViewFront { get; set; } = true;
        public bool ViewRight { get; set; } = true;
        public bool ViewLeft { get; set; } = false;
        public bool ViewBack { get; set; } = false;
        public bool ViewTop { get; set; } = false;

        public bool Tag3D { get; set; } = true;
        public bool TagFront { get; set; } = false;
        public bool TagRight { get; set; } = false;
        public bool TagLeft { get; set; } = false;
        public bool TagBack { get; set; } = false;
        public bool TagTop { get; set; } = false;

        public string OrthoDirection { get; set; } = "NE";

        public string Place3D { get; set; } = "TopLeft";
        public string PlaceBack { get; set; } = "TopRight";
        public string PlaceFront { get; set; } = "BottomLeft";
        public string PlaceLeft { get; set; } = "MidLeft";
        public string PlaceRight { get; set; } = "MidRight";
        public string PlaceTop { get; set; } = "BottomRight";
    }

    public class SpoolOptions
    {
        public string Action { get; set; }                 // "create" | "refresh" | "save-settings"
        public List<long> AssemblyElementIds { get; set; }
        public SpoolSettings Settings { get; set; }
    }

    public class SpoolRequestHandler : IExternalEventHandler
    {
        private UIApplication _uiApp;
        private SpoolOptions _opts;

        public void Set(UIApplication uiApp, SpoolOptions opts, Action<string> _ = null)
        {
            _uiApp = uiApp;
            _opts = opts;
        }

        public void Execute(UIApplication app)
        {
            if (_opts == null) return;
            Document doc = app?.ActiveUIDocument?.Document;
            if (doc == null) return;

            switch (_opts.Action)
            {
                case "save-settings":
                    SaveSettings(doc, _opts.Settings ?? new SpoolSettings());
                    break;

                case "create":
                    using (var tx = new Transaction(doc, "ABMEP – Create Spool Sheets"))
                    {
                        tx.Start();
                        int made = 0;
                        foreach (long id in _opts.AssemblyElementIds ?? new List<long>())
                        {
                            var asm = doc.GetElement(new ElementId(id)) as AssemblyInstance;
                            if (asm == null) continue;
                            CreateSheetAndSchedule(doc, asm, _opts.Settings ?? new SpoolSettings());
                            made++;
                        }
                        tx.Commit();
                        TaskDialog.Show("ABMEP Assembly Manager", $"Created {made} spool sheet(s).");
                    }
                    break;

                case "refresh":
                    using (var tx = new Transaction(doc, "ABMEP – Refresh Schedules on Spool Sheets"))
                    {
                        tx.Start();
                        foreach (long id in _opts.AssemblyElementIds ?? new List<long>())
                        {
                            var asm = doc.GetElement(new ElementId(id)) as AssemblyInstance;
                            if (asm == null) continue;
                            RefreshAssemblySchedule(doc, asm, _opts.Settings ?? new SpoolSettings());
                        }
                        doc.Regenerate();
                        tx.Commit();
                    }
                    break;
            }
        }

        public string GetName() => "ABMEP Spool Request Handler";

        // ---------- settings ----------
        private static void SaveSettings(Document doc, SpoolSettings s)
        {
            var schema = SpoolSettingsStorage.GetOrCreateSchema();
            var ent = doc.ProjectInformation.GetEntity(schema);
            ent = SpoolSettingsStorage.Write(ent, s);
            doc.ProjectInformation.SetEntity(ent);
        }

        // ---------- creation ----------
        private static void CreateSheetAndSchedule(Document doc, AssemblyInstance asm, SpoolSettings s)
        {
            // title block
            var tb = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(fs => string.Equals(fs.Name, s.TitleBlockName, StringComparison.OrdinalIgnoreCase));
            if (tb == null) throw new InvalidOperationException("Selected Title Block not found: " + s.TitleBlockName);

            // sheet (under assembly if supported)
            ViewSheet sheet;
            try { sheet = AssemblyViewUtils.CreateSheet(doc, asm.Id, tb.Id); }
            catch { sheet = ViewSheet.Create(doc, tb.Id); }

            string asmName = GetAssemblyName(doc, asm);
            sheet.Name = asmName;
            sheet.SheetNumber = SuggestUniqueSheetNumber(doc, "SPL-" + asm.Id.Value);

            // create a real Fabrication schedule
            var catId = FabricationCategory.DetectForAssembly(doc, asm);
            var schedule = ViewSchedule.CreateSchedule(doc, catId);
            schedule.Name = $"{asmName} – Schedule";

            // apply template (formatting)
            var tpl = ScheduleFinder.FindTemplateByExactName(doc, s.ScheduleTemplateName);
            if (tpl != null) { try { schedule.ViewTemplateId = tpl.Id; } catch { } }

            // copy columns from best prototype in that category
            var proto = ScheduleFinder.BestPrototypeFor(doc, schedule, s.ScheduleTemplateName);
            if (proto != null) ScheduleFieldSync.CopyFieldsFrom(doc, schedule, proto);

            // filter to this assembly (by name) if a suitable field exists
            AssemblyFilter.ApplyByAssemblyName(schedule, doc, asmName);

            // place 1/16" left, 1/8" top
            SchedulePlacer.Place(doc, sheet, schedule, leftInches: 1.0 / 16.0, topInches: 1.0 / 8.0);
        }

        // ---------- refresh ----------
        private static void RefreshAssemblySchedule(Document doc, AssemblyInstance asm, SpoolSettings s)
        {
            var sheet = FindAssemblySheet(doc, asm);
            if (sheet == null) return;

            // grab any schedule instance on the sheet (non-revision)
            var inst = new FilteredElementCollector(doc, sheet.Id)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>()
                .FirstOrDefault(si => !si.IsTitleblockRevisionSchedule);
            if (inst == null) return;

            var schedule = doc.GetElement(inst.ScheduleId) as ViewSchedule;
            if (schedule == null || schedule.IsTemplate) return;

            // re-apply template
            var tpl = ScheduleFinder.FindTemplateByExactName(doc, s.ScheduleTemplateName);
            if (tpl != null) { try { schedule.ViewTemplateId = tpl.Id; } catch { } }

            // re-sync columns from best prototype in that category
            var proto = ScheduleFinder.BestPrototypeFor(doc, schedule, s.ScheduleTemplateName);
            if (proto != null) ScheduleFieldSync.CopyFieldsFrom(doc, schedule, proto);

            // re-apply filter by assembly name
            AssemblyFilter.ApplyByAssemblyName(schedule, doc, GetAssemblyName(doc, asm));

            // ensure placement stays correct
            SchedulePlacer.Place(doc, sheet, schedule, leftInches: 1.0 / 16.0, topInches: 1.0 / 8.0);
        }

        // ---------- helpers ----------
        private static string GetAssemblyName(Document doc, AssemblyInstance asm)
        {
            var et = doc.GetElement(asm.GetTypeId()) as ElementType;
            return et != null ? et.Name : (asm.Name ?? asm.Id.IntegerValue.ToString());
        }

        private static ViewSheet FindAssemblySheet(Document doc, AssemblyInstance asm)
        {
            string asmName = GetAssemblyName(doc, asm);
            string prefix = "SPL-" + asm.Id.IntegerValue.ToString();

            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .FirstOrDefault(vs =>
                    string.Equals(vs.Name, asmName, StringComparison.OrdinalIgnoreCase) &&
                    vs.SheetNumber != null &&
                    vs.SheetNumber.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }

        private static string SuggestUniqueSheetNumber(Document doc, string baseNum)
        {
            string num = baseNum;
            int i = 1;
            while (SheetNumberExists(doc, num))
            {
                num = baseNum + "-" + i;
                i++;
            }
            return num;
        }

        private static bool SheetNumberExists(Document doc, string num)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Any(vs => string.Equals(vs.SheetNumber, num, StringComparison.OrdinalIgnoreCase));
        }
    }

    public static class SpoolSettingsStorage
    {
        private static readonly Guid SchemaGuidV3 = new Guid("7F6D5A91-6A44-49AF-A4F3-F7B0A8D16A1B");
        private static readonly Guid SchemaGuidV2 = new Guid("E6F0E5A7-0A60-48B9-922E-3F2CF3E5B4E9");

        public static Schema GetOrCreateSchema()
        {
            var s = Schema.Lookup(SchemaGuidV3);
            if (s != null) return s;

            var sb = new SchemaBuilder(SchemaGuidV3);
            sb.SetSchemaName("ABMEP_AssemblyManager_Settings_v3");
            sb.SetReadAccessLevel(AccessLevel.Public);
            sb.SetWriteAccessLevel(AccessLevel.Public);

            sb.AddSimpleField("TitleBlockName", typeof(string));
            sb.AddSimpleField("ScheduleTemplateName", typeof(string));
            sb.AddSimpleField("TagTypeName", typeof(string));
            sb.AddSimpleField("ViewportTypeName", typeof(string));

            sb.AddSimpleField("View3D", typeof(bool));
            sb.AddSimpleField("ViewFront", typeof(bool));
            sb.AddSimpleField("ViewRight", typeof(bool));
            sb.AddSimpleField("ViewLeft", typeof(bool));
            sb.AddSimpleField("ViewBack", typeof(bool));
            sb.AddSimpleField("ViewTop", typeof(bool));

            sb.AddSimpleField("Tag3D", typeof(bool));
            sb.AddSimpleField("TagFront", typeof(bool));
            sb.AddSimpleField("TagRight", typeof(bool));
            sb.AddSimpleField("TagLeft", typeof(bool));
            sb.AddSimpleField("TagBack", typeof(bool));
            sb.AddSimpleField("TagTop", typeof(bool));

            sb.AddSimpleField("OrthoDirection", typeof(string));
            sb.AddSimpleField("Place3D", typeof(string));
            sb.AddSimpleField("PlaceBack", typeof(string));
            sb.AddSimpleField("PlaceFront", typeof(string));
            sb.AddSimpleField("PlaceLeft", typeof(string));
            sb.AddSimpleField("PlaceRight", typeof(string));
            sb.AddSimpleField("PlaceTop", typeof(string));

            return sb.Finish();
        }

        private static T SafeGet<T>(Entity ent, Schema schema, string field, T def)
        {
            try
            {
                var f = schema.GetField(field);
                if (f == null) return def;
                return ent.Get<T>(f);
            }
            catch { return def; }
        }

        private static void SafeSet<T>(Entity ent, Schema schema, string field, T val)
        {
            var f = schema.GetField(field);
            if (f == null) return;
            ent.Set(f, val);
        }

        public static Entity Write(Entity ent, SpoolSettings s)
        {
            var schema = GetOrCreateSchema();
            if (!ent.IsValid()) ent = new Entity(schema);

            SafeSet(ent, schema, "TitleBlockName", s.TitleBlockName ?? "");
            SafeSet(ent, schema, "ScheduleTemplateName", s.ScheduleTemplateName ?? "");
            SafeSet(ent, schema, "TagTypeName", s.TagTypeName ?? "");
            SafeSet(ent, schema, "ViewportTypeName", s.ViewportTypeName ?? "");

            SafeSet(ent, schema, "View3D", s.View3D);
            SafeSet(ent, schema, "ViewFront", s.ViewFront);
            SafeSet(ent, schema, "ViewRight", s.ViewRight);
            SafeSet(ent, schema, "ViewLeft", s.ViewLeft);
            SafeSet(ent, schema, "ViewBack", s.ViewBack);
            SafeSet(ent, schema, "ViewTop", s.ViewTop);

            SafeSet(ent, schema, "Tag3D", s.Tag3D);
            SafeSet(ent, schema, "TagFront", s.TagFront);
            SafeSet(ent, schema, "TagRight", s.TagRight);
            SafeSet(ent, schema, "TagLeft", s.TagLeft);
            SafeSet(ent, schema, "TagBack", s.TagBack);
            SafeSet(ent, schema, "TagTop", s.TagTop);

            SafeSet(ent, schema, "OrthoDirection", s.OrthoDirection ?? "NE");
            SafeSet(ent, schema, "Place3D", s.Place3D ?? "TopLeft");
            SafeSet(ent, schema, "PlaceBack", s.PlaceBack ?? "TopRight");
            SafeSet(ent, schema, "PlaceFront", s.PlaceFront ?? "BottomLeft");
            SafeSet(ent, schema, "PlaceLeft", s.PlaceLeft ?? "MidLeft");
            SafeSet(ent, schema, "PlaceRight", s.PlaceRight ?? "MidRight");
            SafeSet(ent, schema, "PlaceTop", s.PlaceTop ?? "BottomRight");

            return ent;
        }

        public static SpoolSettings Read(Entity ent)
        {
            var v3 = Schema.Lookup(SchemaGuidV3);
            if (v3 != null && ent.Schema != null && ent.Schema.GUID == SchemaGuidV3) return ReadV3(ent, v3);

            var v2 = Schema.Lookup(SchemaGuidV2);
            if (v2 != null && ent.Schema != null && ent.Schema.GUID == SchemaGuidV2) return ReadV2(ent, v2);

            return new SpoolSettings();
        }

        private static SpoolSettings ReadV3(Entity ent, Schema schema)
        {
            var s = new SpoolSettings
            {
                TitleBlockName = SafeGet(ent, schema, "TitleBlockName", ""),
                ScheduleTemplateName = SafeGet(ent, schema, "ScheduleTemplateName", ""),
                TagTypeName = SafeGet(ent, schema, "TagTypeName", ""),
                ViewportTypeName = SafeGet(ent, schema, "ViewportTypeName", ""),
                View3D = SafeGet(ent, schema, "View3D", true),
                ViewFront = SafeGet(ent, schema, "ViewFront", true),
                ViewRight = SafeGet(ent, schema, "ViewRight", true),
                ViewLeft = SafeGet(ent, schema, "ViewLeft", false),
                ViewBack = SafeGet(ent, schema, "ViewBack", false),
                ViewTop = SafeGet(ent, schema, "ViewTop", false),
                Tag3D = SafeGet(ent, schema, "Tag3D", true),
                TagFront = SafeGet(ent, schema, "TagFront", false),
                TagRight = SafeGet(ent, schema, "TagRight", false),
                TagLeft = SafeGet(ent, schema, "TagLeft", false),
                TagBack = SafeGet(ent, schema, "TagBack", false),
                TagTop = SafeGet(ent, schema, "TagTop", false),
                OrthoDirection = SafeGet(ent, schema, "OrthoDirection", "NE"),
                Place3D = SafeGet(ent, schema, "Place3D", "TopLeft"),
                PlaceBack = SafeGet(ent, schema, "PlaceBack", "TopRight"),
                PlaceFront = SafeGet(ent, schema, "PlaceFront", "BottomLeft"),
                PlaceLeft = SafeGet(ent, schema, "PlaceLeft", "MidLeft"),
                PlaceRight = SafeGet(ent, schema, "PlaceRight", "MidRight"),
                PlaceTop = SafeGet(ent, schema, "PlaceTop", "BottomRight")
            };
            return s;
        }

        private static SpoolSettings ReadV2(Entity ent, Schema schema)
        {
            var s = new SpoolSettings
            {
                TitleBlockName = SafeGet(ent, schema, "TitleBlockName", ""),
                ScheduleTemplateName = SafeGet(ent, schema, "ScheduleTemplateName", ""),
                TagTypeName = SafeGet(ent, schema, "TagTypeName", ""),
                ViewportTypeName = SafeGet(ent, schema, "ViewportTypeName", ""),
                View3D = SafeGet(ent, schema, "View3D", true),
                ViewFront = SafeGet(ent, schema, "ViewFront", true),
                ViewRight = SafeGet(ent, schema, "ViewRight", true),
                ViewLeft = SafeGet(ent, schema, "ViewLeft", false),
                ViewBack = SafeGet(ent, schema, "ViewBack", false),
                ViewTop = SafeGet(ent, schema, "ViewTop", false)
            };
            return s;
        }
    }
}
