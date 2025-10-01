// .NET Framework 4.8  |  Platform: AnyCPU/x64
// Purpose: Ribbon host. Builds the "ABMEP" tab from ABMEP.tools.csv placed next to this DLL.
// CSV supports 4, 6, or 8 columns:
//   4 cols: Panel,ButtonLabel,AssemblyPath,FullClassName
//   6 cols: Panel,ButtonLabel,AssemblyPath,FullClassName,SmallImage,LargeImage
//   8 cols: Panel,ButtonLabel,AssemblyPath,FullClassName,SmallImage,LargeImage,Size,StackKey
//     - Size: "small" (stackable) or "large"/blank (normal large slot)
//     - StackKey: groups consecutive small items in the same panel into a stacked column (2–3 items)
//
// Image paths can be absolute or filenames placed next to this DLL.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;   // requires PresentationCore
using Autodesk.Revit.UI;

namespace ABMEP
{
    public class App : IExternalApplication
    {
        private const string TabName = "ABMEP";
        private const string CsvName = "ABMEP.tools.csv";

        public Result OnStartup(UIControlledApplication app)
        {
            try { app.CreateRibbonTab(TabName); } catch { /* already exists */ }

            string csvPath = ResolveCsvPath();
            var rows = LoadRows(csvPath);

            // Preserve panel order as they appear in CSV
            var panelOrder = new List<string>();
            var byPanel = new Dictionary<string, List<Row>>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in rows)
            {
                if (!byPanel.TryGetValue(r.Panel, out var list))
                {
                    list = new List<Row>();
                    byPanel[r.Panel] = list;
                    panelOrder.Add(r.Panel);
                }
                list.Add(r);
            }

            foreach (var panelName in panelOrder)
            {
                RibbonPanel panel;
                try { panel = app.CreateRibbonPanel(TabName, panelName); }
                catch { panel = app.GetRibbonPanels(TabName).FirstOrDefault(p => p.Name.Equals(panelName, StringComparison.OrdinalIgnoreCase)); }

                if (panel == null) continue;

                var items = byPanel[panelName];
                int i = 0;
                while (i < items.Count)
                {
                    var r = items[i];

                    // Handle "small" + StackKey groups
                    if (IsSmall(r) && !string.IsNullOrWhiteSpace(r.StackKey))
                    {
                        // Gather consecutive small rows with the same StackKey
                        var group = new List<Row> { r };
                        int j = i + 1;
                        while (j < items.Count && IsSmall(items[j]) &&
                               string.Equals(items[j].StackKey, r.StackKey, StringComparison.OrdinalIgnoreCase))
                        {
                            group.Add(items[j]);
                            j++;
                        }

                        // Chunk group into stacks of 3 (or 2)
                        int k = 0;
                        while (k < group.Count)
                        {
                            int take = Math.Min(3, group.Count - k);
                            var chunk = group.GetRange(k, take);

                            var pbdList = new List<PushButtonData>(take);
                            foreach (var rr in chunk)
                            {
                                var asmPath = ResolveRelativeToDll(rr.AssemblyPath);
                                if (!File.Exists(asmPath))
                                {
                                    TaskDialog.Show("ABMEP loader", "Assembly not found:\n" + asmPath);
                                    continue;
                                }

                                var pbd = CreatePbd(rr, asmPath);
                                pbdList.Add(pbd);
                            }

                            if (pbdList.Count == 1)
                            {
                                // If we ended up with a single item, just add it as normal.
                                var btn = panel.AddItem(pbdList[0]) as PushButton;
                                ApplyImagesAfterCreate(btn, chunk[0]);
                            }
                            else if (pbdList.Count == 2)
                            {
                                var added = panel.AddStackedItems(pbdList[0], pbdList[1]);
                                ApplyImagesAfterCreate(added, chunk);
                            }
                            else if (pbdList.Count == 3)
                            {
                                var added = panel.AddStackedItems(pbdList[0], pbdList[1], pbdList[2]);
                                ApplyImagesAfterCreate(added, chunk);
                            }

                            k += take;
                        }

                        i = j; // skip past group
                        continue;
                    }

                    // Normal (large) button
                    {
                        var asmPath = ResolveRelativeToDll(r.AssemblyPath);
                        if (!File.Exists(asmPath))
                        {
                            TaskDialog.Show("ABMEP loader", "Assembly not found:\n" + asmPath);
                            i++;
                            continue;
                        }

                        var pbd = CreatePbd(r, asmPath);
                        var item = panel.AddItem(pbd) as PushButton;
                        ApplyImagesAfterCreate(item, r);
                        i++;
                    }
                }
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app) => Result.Succeeded;

        // ---------- helpers ----------

        private static bool IsSmall(Row r)
        {
            var s = (r.Size ?? "").Trim().ToLowerInvariant();
            return s == "small" || s == "s";
        }

        private static PushButtonData CreatePbd(Row r, string asmPath)
        {
            var pbd = new PushButtonData(
                Guid.NewGuid().ToString("N"),
                (r.ButtonLabel ?? "").Replace(@"\n", "\n"),
                asmPath,
                r.FullClassName
            )
            {
                // Set images on the data so stacked buttons get them before creation
                Image = LoadPng(ResolveOptional(r.SmallImage)),
                LargeImage = LoadPng(ResolveOptional(r.LargeImage))
            };
            return pbd;
        }

        private static void ApplyImagesAfterCreate(PushButton btn, Row r)
        {
            if (btn == null) return;
            btn.ToolTip = r.FullClassName;

            // If images weren't set via data (or you want to force), you can still set here
            var small = LoadPng(ResolveOptional(r.SmallImage));
            var large = LoadPng(ResolveOptional(r.LargeImage));
            if (small != null) btn.Image = small;
            if (large != null) btn.LargeImage = large;
        }

        private static void ApplyImagesAfterCreate(IList<RibbonItem> items, List<Row> rows)
        {
            if (items == null || rows == null) return;
            for (int i = 0; i < items.Count && i < rows.Count; i++)
            {
                var btn = items[i] as PushButton;
                if (btn == null) continue;
                ApplyImagesAfterCreate(btn, rows[i]);
            }
        }

        private static string DllDir()
        {
            var loc = typeof(App).Assembly.Location;
            return Path.GetDirectoryName(loc) ?? Environment.CurrentDirectory;
        }

        private static string ResolveCsvPath()
        {
            string nextToDll = Path.Combine(DllDir(), CsvName);
            if (File.Exists(nextToDll)) return nextToDll;

            string roaming = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string addinsBase = Path.Combine(roaming, "Autodesk", "Revit", "Addins");
            string year = Directory.GetDirectories(addinsBase)
                                   .Select(Path.GetFileName)
                                   .Where(n => int.TryParse(n, out _))
                                   .OrderByDescending(n => n)
                                   .FirstOrDefault() ?? "2024";
            string fallback = Path.Combine(addinsBase, year, CsvName);
            return File.Exists(fallback) ? fallback : nextToDll;
        }

        private static string ResolveRelativeToDll(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return path ?? "";
            return Path.IsPathRooted(path) ? path : Path.Combine(DllDir(), path);
        }

        private static string ResolveOptional(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "";
            return Path.IsPathRooted(path) ? path : Path.Combine(DllDir(), path);
        }

        private static ImageSource LoadPng(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
            try
            {
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.UriSource = new Uri(path, UriKind.Absolute);
                bi.EndInit();
                bi.Freeze();           // Revit UI thread safety
                return bi;
            }
            catch { return null; }
        }

        private static List<Row> LoadRows(string csvPath)
        {
            var list = new List<Row>();
            if (!File.Exists(csvPath)) return list;

            foreach (var raw in File.ReadAllLines(csvPath))
            {
                var line = raw?.Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;
                if (line.StartsWith("#")) continue;

                // parse up to 8 parts; shorter lines are allowed
                var parts = SplitCsv(line, 8);
                if (parts.Length < 4) continue;

                var row = new Row
                {
                    Panel = Unquote(parts[0]),
                    ButtonLabel = Unquote(parts[1]),
                    AssemblyPath = Unquote(parts[2]),
                    FullClassName = Unquote(parts[3]),
                    SmallImage = parts.Length >= 5 ? Unquote(parts[4]) : null,
                    LargeImage = parts.Length >= 6 ? Unquote(parts[5]) : null,
                    Size = parts.Length >= 7 ? Unquote(parts[6]) : null,
                    StackKey = parts.Length >= 8 ? Unquote(parts[7]) : null
                };
                list.Add(row);
            }
            return list;
        }

        private static string[] SplitCsv(string line, int maxParts)
        {
            var result = new List<string>(maxParts);
            bool inQuotes = false;
            var cur = new System.Text.StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (ch == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    { cur.Append('"'); i++; }
                    else { inQuotes = !inQuotes; }
                }
                else if (ch == ',' && !inQuotes)
                { result.Add(cur.ToString()); cur.Clear(); }
                else
                { cur.Append(ch); }
            }
            result.Add(cur.ToString());
            while (result.Count < maxParts) result.Add(string.Empty);
            return result.ToArray();
        }

        private static string Unquote(string s)
        {
            if (string.IsNullOrEmpty(s)) return s ?? "";
            s = s.Trim();
            if (s.Length >= 2 && s.StartsWith("\"") && s.EndsWith("\""))
                s = s.Substring(1, s.Length - 2).Replace("\"\"", "\"");
            return s;
        }

        private class Row
        {
            public string Panel { get; set; }
            public string ButtonLabel { get; set; }
            public string AssemblyPath { get; set; }
            public string FullClassName { get; set; }
            public string SmallImage { get; set; }   // optional (16x16)
            public string LargeImage { get; set; }   // optional (32x32)
            public string Size { get; set; }         // optional ("small" stacks)
            public string StackKey { get; set; }     // optional (group consecutive small items)
        }
    }
}
