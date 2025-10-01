using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace ABMEP.Work.Views
{
    public partial class AssemblySettingsWindow : Window
    {
        private readonly SpoolSettings _initial;
        public SpoolSettings Result { get; private set; }

        private readonly List<string> _titleBlocks;
        private readonly List<string> _schedules;
        private readonly List<string> _tagTypes;
        private readonly List<string> _viewportTypes;

        private static readonly string[] PlacementOptions = new[]
        {
            "TopLeft","TopCenter","TopRight",
            "MidLeft","MidCenter","MidRight",
            "BottomLeft","BottomCenter","BottomRight"
        };

        // --- Persist window size/position across sessions ---
        private const string UiFileName = "AssemblySettingsWindow.ui";
        private static string UiFilePath
        {
            get
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ABMEP");
                Directory.CreateDirectory(dir);
                return Path.Combine(dir, UiFileName);
            }
        }

        private void LoadWindowBounds()
        {
            try
            {
                if (!File.Exists(UiFilePath)) return;
                var s = File.ReadAllText(UiFilePath).Trim(); // W;H;L;T;State
                var p = s.Split(';');
                if (p.Length >= 2)
                {
                    if (double.TryParse(p[0], out var w) && w > 200) Width = w;
                    if (double.TryParse(p[1], out var h) && h > 200) Height = h;
                }
                if (p.Length >= 4)
                {
                    if (double.TryParse(p[2], out var l)) Left = l;
                    if (double.TryParse(p[3], out var t)) Top = t;
                }
                if (p.Length >= 5 && Enum.TryParse(p[4], out WindowState st))
                    WindowState = st;
            }
            catch { /* ignore */ }
        }

        private void SaveWindowBounds()
        {
            try
            {
                // Save “normal” size if currently maximized
                var r = WindowState == WindowState.Normal ? new Rect(Left, Top, Width, Height) : RestoreBounds;
                var payload = string.Join(";", new[]
                {
                    r.Width.ToString("0.##"),
                    r.Height.ToString("0.##"),
                    r.Left.ToString("0.##"),
                    r.Top.ToString("0.##"),
                    WindowState.ToString()
                });
                File.WriteAllText(UiFilePath, payload);
            }
            catch { /* ignore */ }
        }

        public AssemblySettingsWindow(
            SpoolSettings settings,
            List<string> titleBlocks,
            List<string> scheduleTemplates,
            List<string> tagTypes,
            List<string> viewportTypes,
            string logoPath)
        {
            InitializeComponent();

            // remember/restore window size & position
            Loaded += (_, __) => LoadWindowBounds();
            Closing += (_, __) => SaveWindowBounds();

            _initial = settings ?? new SpoolSettings();
            _titleBlocks = titleBlocks ?? new List<string>();
            _schedules = scheduleTemplates ?? new List<string>();
            _tagTypes = tagTypes ?? new List<string>();
            _viewportTypes = viewportTypes ?? new List<string>();

            // logo
            TryLoadLogo(logoPath);

            // top pickers
            cmbTitleBlock.ItemsSource = _titleBlocks;
            cmbScheduleTemplate.ItemsSource = _schedules;
            cmbTagType.ItemsSource = _tagTypes;
            cmbViewportType.ItemsSource = _viewportTypes;

            cmbTitleBlock.SelectedItem = pick(_initial.TitleBlockName, _titleBlocks);
            cmbScheduleTemplate.SelectedItem = pick(_initial.ScheduleTemplateName, _schedules);
            cmbTagType.SelectedItem = pick(_initial.TagTypeName, _tagTypes);
            cmbViewportType.SelectedItem = pick(_initial.ViewportTypeName, _viewportTypes);

            // direction + placements
            cmb3DDirection.SelectedIndex = dirIndex(_initial.OrthoDirection);
            foreach (var cb in new[] { cmb3DPlace, cmbBackPlace, cmbFrontPlace, cmbLeftPlace, cmbRightPlace, cmbTopPlace })
                cb.ItemsSource = PlacementOptions;

            cmb3DPlace.SelectedItem = pick(_initial.Place3D, PlacementOptions);
            cmbBackPlace.SelectedItem = pick(_initial.PlaceBack, PlacementOptions);
            cmbFrontPlace.SelectedItem = pick(_initial.PlaceFront, PlacementOptions);
            cmbLeftPlace.SelectedItem = pick(_initial.PlaceLeft, PlacementOptions);
            cmbRightPlace.SelectedItem = pick(_initial.PlaceRight, PlacementOptions);
            cmbTopPlace.SelectedItem = pick(_initial.PlaceTop, PlacementOptions);

            // include + tag flags
            chk3D.IsChecked = _initial.View3D;
            chkFront.IsChecked = _initial.ViewFront;
            chkRight.IsChecked = _initial.ViewRight;
            chkLeft.IsChecked = _initial.ViewLeft;
            chkBack.IsChecked = _initial.ViewBack;
            chkTop.IsChecked = _initial.ViewTop;

            chk3DTag.IsChecked = _initial.Tag3D;
            chkFrontTag.IsChecked = _initial.TagFront;
            chkRightTag.IsChecked = _initial.TagRight;
            chkLeftTag.IsChecked = _initial.TagLeft;
            chkBackTag.IsChecked = _initial.TagBack;
            chkTopTag.IsChecked = _initial.TagTop;

            // helpers
            string pick(string name, IList<string> list) =>
                list?.FirstOrDefault(x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase))
                ?? (list?.FirstOrDefault() ?? null);

            int dirIndex(string d)
            {
                var all = new[] { "NE", "NW", "SE", "SW", "N", "S", "E", "W", "ISO" };
                int i = Array.FindIndex(all, x => string.Equals(x, d, StringComparison.OrdinalIgnoreCase));
                return i < 0 ? 0 : i;
            }
        }

        private void TryLoadLogo(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.UriSource = new Uri(path, UriKind.Absolute);
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    imgLogo.Source = bmp;
                }
                else
                {
                    imgLogo.Visibility = Visibility.Collapsed;
                }
            }
            catch { imgLogo.Visibility = Visibility.Collapsed; }
        }

        private string Get3DDirection()
        {
            if (cmb3DDirection.SelectedItem is ComboBoxItem ci)
                return ci.Content?.ToString() ?? "NE";
            return "NE";
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            var s = _initial ?? new SpoolSettings();

            s.TitleBlockName = (string)cmbTitleBlock.SelectedItem ?? s.TitleBlockName;
            s.ScheduleTemplateName = (string)cmbScheduleTemplate.SelectedItem ?? s.ScheduleTemplateName;
            s.TagTypeName = (string)cmbTagType.SelectedItem ?? s.TagTypeName;
            s.ViewportTypeName = (string)cmbViewportType.SelectedItem ?? s.ViewportTypeName;

            s.View3D = chk3D.IsChecked == true;
            s.ViewFront = chkFront.IsChecked == true;
            s.ViewRight = chkRight.IsChecked == true;
            s.ViewLeft = chkLeft.IsChecked == true;
            s.ViewBack = chkBack.IsChecked == true;
            s.ViewTop = chkTop.IsChecked == true;

            s.Tag3D = chk3DTag.IsChecked == true;
            s.TagFront = chkFrontTag.IsChecked == true;
            s.TagRight = chkRightTag.IsChecked == true;
            s.TagLeft = chkLeftTag.IsChecked == true;
            s.TagBack = chkBackTag.IsChecked == true;
            s.TagTop = chkTopTag.IsChecked == true;

            s.OrthoDirection = Get3DDirection();

            s.Place3D = (string)cmb3DPlace.SelectedItem ?? s.Place3D;
            s.PlaceBack = (string)cmbBackPlace.SelectedItem ?? s.PlaceBack;
            s.PlaceFront = (string)cmbFrontPlace.SelectedItem ?? s.PlaceFront;
            s.PlaceLeft = (string)cmbLeftPlace.SelectedItem ?? s.PlaceLeft;
            s.PlaceRight = (string)cmbRightPlace.SelectedItem ?? s.PlaceRight;
            s.PlaceTop = (string)cmbTopPlace.SelectedItem ?? s.PlaceTop;

            Result = s;
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
