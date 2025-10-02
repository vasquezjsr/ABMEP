// Target: .NET Framework 4.8
// Assembly: ABMEP.Work.dll

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
// Alias WinForms to avoid 'View' ambiguity with Revit
using WinForms = System.Windows.Forms;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Fabrication;
using Autodesk.Revit.UI;

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class HangerInsertsCommand : IExternalCommand
    {
        private const string OUTPUT_DIR = @"C:\Temp";

        // ===== Entry =====
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = data.Application.ActiveUIDocument;
                Document doc = uidoc?.Document;
                if (doc == null || doc.ActiveView == null)
                {
                    message = "No active document/view.";
                    return Result.Failed;
                }

                var view = doc.ActiveView;

                // Collect hangers in ACTIVE VIEW
                var hangers = new FilteredElementCollector(doc, view.Id)
                    .OfCategory(BuiltInCategory.OST_FabricationHangers)
                    .WhereElementIsNotElementType()
                    .OfType<FabricationPart>()
                    .ToList();

                if (hangers.Count == 0)
                {
                    TaskDialog.Show("Hanger Inserts", "No MEP Fabrication Hangers found in the active view.");
                    return Result.Succeeded;
                }

                // Level prompt (for filename label)
                string levelDefault = GetDefaultLevelForView(view);
                string levelText = PromptForLevel(levelDefault);

                // Count map: Size -> Count
                var countsBySize = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                int skippedNoSize = 0;   // length present but no readable diameter
                int totalRods = 0;

                foreach (var fp in hangers)
                {
                    // Read ordered rod diameters (rod_1, rod_2, ...). We only need first 2 for A/B.
                    var rodFeet = GetRodDiametersFeetOrdered(fp);
                    var rodSizes = rodFeet.Select(f => InchesToNiceFraction(f * 12.0))
                                          .Where(s => !string.IsNullOrWhiteSpace(s))
                                          .ToList();

                    // Read eVolve lengths
                    bool hasA = HasNonZeroLength(GetParamByNameVariants(fp, "eM_Length A", "eM_LengthA"));
                    bool hasB = HasNonZeroLength(GetParamByNameVariants(fp, "eM_Length B", "eM_LengthB"));

                    // Nothing present on this hanger → 0 rods
                    if (!hasA && !hasB) continue;

                    // We need at least one recognizable diameter to place counts
                    if (rodSizes.Count == 0)
                    {
                        skippedNoSize += (hasA ? 1 : 0) + (hasB ? 1 : 0);
                        continue;
                    }

                    // Map A -> rodSizes[0]
                    if (hasA)
                    {
                        string sA = rodSizes[0];
                        Add(countsBySize, sA, 1);
                        totalRods++;
                    }

                    // Map B -> rodSizes[1] if present; else fall back to rodSizes[0]
                    if (hasB)
                    {
                        string sB = rodSizes.Count >= 2 ? rodSizes[1] : rodSizes[0];
                        Add(countsBySize, sB, 1);
                        totalRods++;
                    }
                }

                if (totalRods == 0)
                {
                    TaskDialog.Show("Hanger Inserts",
                        "No rods detected from eVolve 'eM_Length A/B' in the active view.");
                    return Result.Succeeded;
                }

                // Build filename
                string proj = GetProjectName(doc);
                string date = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                string fileName = $"{San(proj)}_{San(levelText)}_Hanger Inserts_{date}.csv";
                Directory.CreateDirectory(OUTPUT_DIR);
                string path = Path.Combine(OUTPUT_DIR, fileName);

                // Sort sizes numerically by inches where possible
                var rows = countsBySize
                    .Select(kv => new
                    {
                        Size = kv.Key,
                        Inches = TryParseInchesText(kv.Key, out double ins) ? ins : double.PositiveInfinity,
                        Cnt = kv.Value
                    })
                    .OrderBy(x => x.Inches)
                    .ThenBy(x => x.Size, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                using (var sw = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
                {
                    sw.WriteLine("Project,Level,Size,Count");
                    foreach (var r in rows)
                    {
                        sw.WriteLine(string.Join(",",
                            Csv(proj),
                            Csv(levelText),
                            Csv(r.Size),
                            r.Cnt.ToString(CultureInfo.InvariantCulture)
                        ));
                    }
                }

                TaskDialog.Show("Hanger Inserts",
                    $"Exported insert counts to:\n{path}\n\n" +
                    $"Rod sizes: {countsBySize.Count}\n" +
                    $"Total inserts (rods): {totalRods}" +
                    (skippedNoSize > 0 ? $"\nRods skipped (no readable size): {skippedNoSize}" : "")
                );

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
        }

        // ===== eVolve helpers =====

        private static Parameter GetParamByNameVariants(Element e, params string[] names)
        {
            if (e == null || names == null || names.Length == 0) return null;

            // Normalize function: lower + remove spaces, underscores, hyphens
            string Norm(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return "";
                var sb = new StringBuilder(s.Length);
                foreach (char c in s.ToLowerInvariant())
                {
                    if (c != ' ' && c != '_' && c != '-') sb.Append(c);
                }
                return sb.ToString();
            }

            var want = new HashSet<string>(names.Select(Norm));

            foreach (Parameter p in e.Parameters)
            {
                string n = p?.Definition?.Name ?? "";
                if (want.Contains(Norm(n))) return p;
            }
            return null;
        }

        private static bool HasNonZeroLength(Parameter p)
        {
            if (p == null || !p.HasValue) return false;

            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        return p.AsDouble() > 1e-8; // internal feet
                    case StorageType.Integer:
                        return p.AsInteger() != 0;
                    case StorageType.String:
                        {
                            var s = p.AsString();
                            if (string.IsNullOrWhiteSpace(s)) return false;
                            // Try parse feet/inch or raw number
                            if (TryParseFeetInchesString(s, out double ft)) return ft > 1e-8;
                            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                                return Math.Abs(v) > 1e-8;
                            return true; // non-empty string → treat as present
                        }
                    default:
                        try
                        {
                            string vs = p.AsValueString();
                            if (string.IsNullOrWhiteSpace(vs)) return false;
                            if (TryParseFeetInchesString(vs, out double ft)) return ft > 1e-8;
                            if (double.TryParse(vs, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                                return Math.Abs(v) > 1e-8;
                            return true;
                        }
                        catch { return true; }
                }
            }
            catch { return false; }
        }

        // ===== Diameter reading (params only) =====

        private static List<double> GetRodDiametersFeetOrdered(FabricationPart fp)
        {
            var byIndex = new SortedDictionary<int, double>();
            var loose = new List<double>();

            try
            {
                foreach (Parameter p in fp.Parameters)
                {
                    if (p == null || !p.HasValue) continue;
                    string def = p.Definition != null ? p.Definition.Name : "";
                    if (string.IsNullOrWhiteSpace(def)) continue;

                    string low = def.ToLowerInvariant();
                    if (!(low.Contains("rod") && low.Contains("diameter"))) continue;

                    double valFeet = double.NaN;

                    if (p.StorageType == StorageType.Double)
                    {
                        valFeet = p.AsDouble(); // internal feet
                    }
                    else
                    {
                        string s = null;
                        try { s = p.AsValueString(); } catch { }
                        if (string.IsNullOrWhiteSpace(s) && p.StorageType == StorageType.String)
                            s = p.AsString();

                        if (!string.IsNullOrWhiteSpace(s) && TryParseFeetInchesString(s, out double vf))
                            valFeet = vf;
                    }

                    if (!IsFinite(valFeet) || valFeet <= 0) continue;

                    int idx = ExtractRodIndex(low);
                    if (idx > 0) byIndex[idx] = valFeet;
                    else loose.Add(valFeet);
                }
            }
            catch { }

            var list = new List<double>();
            foreach (var kv in byIndex) list.Add(kv.Value);
            list.AddRange(loose);
            return list;
        }

        private static int ExtractRodIndex(string nameLower)
        {
            for (int i = 1; i <= 8; i++)
            {
                if (nameLower.Contains("rod_" + i) || nameLower.Contains("rod " + i) || nameLower.Contains("rod-" + i))
                    return i;
            }
            return -1;
        }

        // ===== Formatting / parsing =====

        private static bool TryParseFeetInchesString(string s, out double feet)
        {
            feet = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            try
            {
                string t = s.Trim()
                            .Replace("feet", "'").Replace("foot", "'").Replace("ft", "'")
                            .Replace("inches", "\"").Replace("inch", "\"").Replace(" in", "\"").Replace("in ", "\" ");

                int fidx = t.IndexOf("'");
                double ft = 0, inch = 0;

                if (fidx >= 0)
                {
                    double.TryParse(t.Substring(0, fidx).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out ft);
                    t = t.Substring(fidx + 1).Trim();
                }

                t = t.Replace("\"", "").Replace("-", " ");
                var parts = t.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    if (part.Contains("/"))
                    {
                        var ab = part.Split('/');
                        if (ab.Length == 2 &&
                            double.TryParse(ab[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double a) &&
                            double.TryParse(ab[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double b) &&
                            b != 0) inch += a / b;
                    }
                    else if (double.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                    {
                        inch += v;
                    }
                }

                // crude mm support
                if (s.ToLowerInvariant().Contains("mm") &&
                    double.TryParse(s.ToLowerInvariant().Replace("mm", "").Trim(),
                        NumberStyles.Float, CultureInfo.InvariantCulture, out double mm))
                {
                    feet = mm / 304.8;
                    return true;
                }

                feet = ft + inch / 12.0;
                return (ft != 0 || inch != 0);
            }
            catch { feet = 0; return false; }
        }

        private static bool TryParseInchesText(string sizeText, out double inches)
        {
            inches = 0;
            if (string.IsNullOrWhiteSpace(sizeText)) return false;
            try
            {
                string s = sizeText.Trim().Replace("\"", "");
                double total = 0;
                foreach (var part in s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (part.Contains("/"))
                    {
                        var ab = part.Split('/');
                        if (ab.Length == 2 &&
                            double.TryParse(ab[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double a) &&
                            double.TryParse(ab[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double b) &&
                            b != 0)
                        {
                            total += a / b;
                        }
                    }
                    else if (double.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
                    {
                        total += v;
                    }
                }
                inches = total;
                return total > 0;
            }
            catch { return false; }
        }

        private static string InchesToNiceFraction(double inches)
        {
            double frac = Math.Round(inches * 16.0) / 16.0; // nearest 1/16"
            int whole = (int)Math.Floor(frac);
            double rem = frac - whole;

            int num = (int)Math.Round(rem * 16.0);
            int den = 16;

            if (num == 0) return whole.ToString(CultureInfo.InvariantCulture) + "\"";

            int g = GCD(num, den); num /= g; den /= g;

            if (whole == 0)
                return num.ToString(CultureInfo.InvariantCulture) + "/" + den.ToString(CultureInfo.InvariantCulture) + "\"";
            else
                return whole.ToString(CultureInfo.InvariantCulture) + " " + num.ToString(CultureInfo.InvariantCulture) + "/" + den.ToString(CultureInfo.InvariantCulture) + "\"";
        }

        private static int GCD(int a, int b)
        {
            a = Math.Abs(a); b = Math.Abs(b);
            while (b != 0) { int t = a % b; a = b; b = t; }
            return a == 0 ? 1 : a;
        }

        private static void Add(Dictionary<string, int> map, string key, int inc)
        {
            if (string.IsNullOrWhiteSpace(key)) return;
            if (!map.ContainsKey(key)) map[key] = 0;
            map[key] += inc;
        }

        // ===== Project/View/UI helpers =====

        private static string GetProjectName(Document doc)
        {
            string name = doc.ProjectInformation?.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = !string.IsNullOrWhiteSpace(doc.Title)
                    ? doc.Title
                    : Path.GetFileNameWithoutExtension(doc.PathName);
            }
            return string.IsNullOrWhiteSpace(name) ? "Project" : name.Trim();
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

        private static string San(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Untitled";
            var bad = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name) sb.Append(Array.IndexOf(bad, ch) >= 0 ? '_' : ch);
            return sb.ToString().Trim();
        }

        private static string Csv(string s)
        {
            if (s == null) s = "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static string PromptForLevel(string defaultLevel)
        {
            using (var form = new LevelPromptForm(defaultLevel ?? "Level"))
            {
                var result = form.ShowDialog();
                if (result == WinForms.DialogResult.OK)
                {
                    var text = (form.LevelText ?? "").Trim();
                    return string.IsNullOrEmpty(text) ? (defaultLevel ?? "Level") : text;
                }
                return defaultLevel ?? "Level";
            }
        }

        private class LevelPromptForm : WinForms.Form
        {
            private readonly WinForms.TextBox _tb;
            public string LevelText => _tb.Text;

            public LevelPromptForm(string defaultLevel)
            {
                this.Text = "Hanger Inserts – Level";
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

        private static bool IsFinite(double v) => !(double.IsNaN(v) || double.IsInfinity(v));
    }
}
