// ABMEP.Work/SleevesTrimbleCommand.cs
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
// Alias WinForms to avoid 'View' ambiguity with Revit
using WinForms = System.Windows.Forms;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class SleevesTrimbleCommand : IExternalCommand
    {
        private const string Folder = @"C:\Temp";

        // Parameter names for Layer
        private const string RunNominalDiameterParam = "Run_Nominal_Diameter";
        private const string ServiceAbbrevParam = "eM_Service Abbreviation";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                if (uiDoc == null || uiDoc.Document == null)
                {
                    message = "No active document.";
                    return Result.Failed;
                }

                Document doc = uiDoc.Document;
                Autodesk.Revit.DB.View view = doc.ActiveView;

                // Collect sleeves in the ACTIVE VIEW
                var sleeves = CollectSleevesInView(doc, view);
                if (!sleeves.Any())
                {
                    TaskDialog.Show("Sleeves Trimble", "No sleeves found in the active view.");
                    return Result.Succeeded;
                }

                // Build records once so both files are consistent
                var records = BuildRecords(sleeves);

                // Build filename parts
                string projectName = GetProjectName(doc);
                string defaultLevel = GetDefaultLevelForView(view);
                string level = PromptForLevel(defaultLevel);
                string dateStr = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                string safeProject = SafeFilePart(projectName);
                string safeLevel = SafeFilePart(level);

                string baseName = $"{safeProject}_{safeLevel}_Sleeves Trimble_{dateStr}";
                string pathBase = Path.Combine(Folder, baseName + ".csv");
                string pathWithIds = Path.Combine(Folder, baseName + "_WithIds.csv");

                if (!Directory.Exists(Folder)) Directory.CreateDirectory(Folder);

                // =========================
                // File 1: standard export
                // =========================
                using (var sw = new StreamWriter(pathBase, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
                {
                    sw.WriteLine("Name,X,Y,Z,Description"); // NEW headers
                    foreach (var r in records)
                    {
                        string line = string.Join(",",
                            Csv(r.PointId),      // Name
                            Num(r.X_ft),         // X
                            Num(r.Y_ft),         // Y
                            Num(r.Z_ft),         // Z
                            Csv(r.Layer)         // Description (Layer only)
                        );
                        sw.WriteLine(line);
                    }
                }

                // ==================================
                // File 2: includes ElementId in Description
                // ==================================
                using (var sw = new StreamWriter(pathWithIds, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
                {
                    sw.WriteLine("Name,X,Y,Z,Description"); // SAME headers
                    foreach (var r in records)
                    {
                        string desc = string.IsNullOrEmpty(r.Layer)
                            ? $"ID {r.ElementIdText}"
                            : $"{r.Layer} | ID {r.ElementIdText}";

                        string line = string.Join(",",
                            Csv(r.PointId),      // Name
                            Num(r.X_ft),         // X
                            Num(r.Y_ft),         // Y
                            Num(r.Z_ft),         // Z
                            Csv(desc)            // Description (Layer + ID)
                        );
                        sw.WriteLine(line);
                    }
                }

                TaskDialog.Show("Sleeves Trimble",
                    $"Exported {records.Count} sleeves to:\n{pathBase}\n{pathWithIds}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // ===== Records =====
        private class SleeveRecord
        {
            public string PointId { get; set; }
            public double X_ft { get; set; }
            public double Y_ft { get; set; }
            public double Z_ft { get; set; }
            public string Layer { get; set; }
            public long ElementId { get; set; } // Revit 2024+: ElementId.Value is long
            public string ElementIdText => ElementId.ToString(CultureInfo.InvariantCulture);
        }

        private static List<SleeveRecord> BuildRecords(IList<Element> sleeves)
        {
            var list = new List<SleeveRecord>(sleeves.Count);
            int idx = 1;

            foreach (var e in sleeves)
            {
                XYZ center = GetElementCenter(e);
                double topZ = GetTopElevationFeet(e);
                string layer = BuildLayer(e);

                list.Add(new SleeveRecord
                {
                    PointId = $"P-{idx}",
                    X_ft = center.X,
                    Y_ft = center.Y,
                    Z_ft = topZ,
                    Layer = layer,
                    ElementId = e.Id.Value
                });

                idx++;
            }

            return list;
        }

        // ===== Collection / Geometry =====
        private static IList<Element> CollectSleevesInView(Document doc, Autodesk.Revit.DB.View view)
        {
            var cats = new HashSet<BuiltInCategory>
            {
                BuiltInCategory.OST_ConduitFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_GenericModel
            };

            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && cats.Contains((BuiltInCategory)e.Category.Id.IntegerValue));

            // Heuristic: family/type/name contains "sleeve"
            var sleeves = collector.Where(e =>
            {
                string fam = e is FamilyInstance fi && fi.Symbol != null ? fi.Symbol.FamilyName : e.Name;
                string type = e is FamilyInstance fi2 && fi2.Symbol != null ? fi2.Symbol.Name : string.Empty;
                return (fam ?? string.Empty).IndexOf("sleeve", StringComparison.OrdinalIgnoreCase) >= 0
                    || (type ?? string.Empty).IndexOf("sleeve", StringComparison.OrdinalIgnoreCase) >= 0
                    || (e.Name ?? string.Empty).IndexOf("sleeve", StringComparison.OrdinalIgnoreCase) >= 0;
            }).ToList();

            return sleeves;
        }

        private static XYZ GetElementCenter(Element e)
        {
            if (e.Location is LocationPoint lp)
                return lp.Point;

            var bb = e.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) * 0.5;

            return XYZ.Zero;
        }

        private static double GetTopElevationFeet(Element e)
        {
            var bb = e.get_BoundingBox(null);
            if (bb != null) return bb.Max.Z; // internal units are feet
            return GetElementCenter(e).Z;
        }

        // ===== Layer formatting =====
        private static string BuildLayer(Element e)
        {
            string diaInches = GetRunNominalDiameterInchesText(e); // "8", "4", etc.
            string svc = GetParamAsString(e, ServiceAbbrevParam)?.Trim() ?? "";

            if (string.IsNullOrEmpty(diaInches) && string.IsNullOrEmpty(svc)) return "";
            if (string.IsNullOrEmpty(diaInches)) return svc;
            if (string.IsNullOrEmpty(svc)) return diaInches;
            return $"{diaInches} {svc}";
        }

        private static string GetRunNominalDiameterInchesText(Element e)
        {
            var p = e.LookupParameter(RunNominalDiameterParam);
            if (p == null) return "";

            if (p.StorageType == StorageType.Double)
            {
                double feet = p.AsDouble();
                double inches = Math.Round(feet * 12.0, MidpointRounding.AwayFromZero);
                return inches <= 0 ? "" : inches.ToString("0", CultureInfo.InvariantCulture);
            }

            try
            {
                string vs = p.AsValueString();
                string parsed = ParseInchesFromAnyText(vs);
                if (!string.IsNullOrEmpty(parsed)) return parsed;
            }
            catch { /* ignore */ }

            if (p.StorageType == StorageType.String)
            {
                string s = p.AsString();
                string parsed = ParseInchesFromAnyText(s);
                if (!string.IsNullOrEmpty(parsed)) return parsed;
            }

            return "";
        }

        private static string ParseInchesFromAnyText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            var m = Regex.Match(text, @"([-+]?\d*\.?\d+)");
            if (!m.Success) return "";
            if (!double.TryParse(m.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
            {
                if (!double.TryParse(m.Value, out val)) return "";
            }
            return Math.Round(val, MidpointRounding.AwayFromZero).ToString("0", CultureInfo.InvariantCulture);
        }

        private static string GetParamAsString(Element e, string paramName)
        {
            Parameter p = e.LookupParameter(paramName);
            if (p == null) return null;

            switch (p.StorageType)
            {
                case StorageType.String:
                    return p.AsString();
                case StorageType.Integer:
                    return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                case StorageType.Double:
                    return p.AsDouble().ToString(CultureInfo.InvariantCulture);
                case StorageType.ElementId:
                    return p.AsElementId().Value.ToString(CultureInfo.InvariantCulture);
                default:
                    return null;
            }
        }

        // ===== Filename helpers =====
        private static string GetProjectName(Document doc)
        {
            string name = doc.ProjectInformation?.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = !string.IsNullOrWhiteSpace(doc.Title)
                    ? doc.Title
                    : Path.GetFileNameWithoutExtension(doc.PathName);
            }
            return string.IsNullOrWhiteSpace(name) ? "Project" : name;
        }

        private static string GetDefaultLevelForView(Autodesk.Revit.DB.View view)
        {
            try
            {
                Level lvl = view.GenLevel;
                if (lvl != null && !string.IsNullOrWhiteSpace(lvl.Name))
                    return lvl.Name;
            }
            catch { /* some views don't support GenLevel */ }

            return view?.Name ?? "Level";
        }

        private static string PromptForLevel(string defaultLevel)
        {
            using (var form = new LevelPromptForm(defaultLevel ?? "Level"))
            {
                WinForms.DialogResult result = form.ShowDialog();
                if (result == WinForms.DialogResult.OK)
                {
                    var text = (form.LevelText ?? "").Trim();
                    return string.IsNullOrEmpty(text) ? (defaultLevel ?? "Level") : text;
                }
                return defaultLevel ?? "Level";
            }
        }

        private static string SafeFilePart(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "Untitled";
            var invalid = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                sb.Append(invalid.Contains(c) ? '_' : c);
            }
            return sb.ToString().Trim();
        }

        // ===== CSV helpers =====
        private static string Csv(string s)
        {
            if (s == null) s = "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static string Num(double feet)
        {
            return feet.ToString("0.###", CultureInfo.InvariantCulture);
        }

        // ===== Small inline WinForms prompt =====
        private class LevelPromptForm : WinForms.Form
        {
            private readonly WinForms.TextBox _tb;
            public string LevelText => _tb.Text;

            public LevelPromptForm(string defaultLevel)
            {
                this.Text = "Sleeves Trimble – Level";
                this.StartPosition = WinForms.FormStartPosition.CenterScreen;
                this.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = false;
                this.Width = 420;
                this.Height = 160;

                var lbl = new WinForms.Label
                {
                    Text = "Enter Level/Area label for the export filename:",
                    AutoSize = true,
                    Left = 12,
                    Top = 12
                };

                _tb = new WinForms.TextBox
                {
                    Left = 15,
                    Top = 40,
                    Width = 370,
                    Text = defaultLevel ?? "Level"
                };

                var btnOk = new WinForms.Button
                {
                    Text = "OK",
                    DialogResult = WinForms.DialogResult.OK,
                    Left = 220,
                    Top = 80,
                    Width = 75
                };

                var btnCancel = new WinForms.Button
                {
                    Text = "Cancel",
                    DialogResult = WinForms.DialogResult.Cancel,
                    Left = 305,
                    Top = 80,
                    Width = 75
                };

                this.Controls.Add(lbl);
                this.Controls.Add(_tb);
                this.Controls.Add(btnOk);
                this.Controls.Add(btnCancel);

                this.AcceptButton = btnOk;
                this.CancelButton = btnCancel;
            }
        }
    }
}
