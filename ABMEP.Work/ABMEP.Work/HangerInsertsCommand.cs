// Target: .NET Framework 4.8
// Assembly: ABMEP.Work.dll

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Fabrication;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class HangerInsertsCommand : IExternalCommand
    {
        // ---- parameter name pools ----
        private static readonly string[] RodCountParamNames = {
            "Number of Rods","Rod Quantity","Rod Count","Support Rods","Rods"
        };

        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            UIDocument uidoc = data.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            if (doc == null || doc.ActiveView == null)
            {
                message = "No active document/view.";
                return Result.Cancelled;
            }

            // collect hangers in the ACTIVE VIEW (matches your Trimble tool scope)
            var hangers = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_FabricationHangers)
                .WhereElementIsNotElementType()
                .OfType<FabricationPart>()
                .ToList();

            if (hangers.Count == 0)
            {
                TaskDialog.Show("ABMEP - Hanger Inserts", "No MEP Fabrication Hangers found in the active view.");
                return Result.Cancelled;
            }

            // prompt for Level TEXT (same UX concept as HangersTrimbleCommand)
            string levelText = PromptLevel("Level");

            // tally (Insert Size -> Count of inserts)
            var perSize = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int grandTotal = 0;

            foreach (var fp in hangers)
            {
                // 1) get ordered rod diameters as nice strings (3/8", 1/2", etc.)
                var diaStrings = GetRodDiameterStrings(fp); // uses same logic as Trimble
                string primary = (diaStrings.Count > 0 ? diaStrings[0] : "Unknown");

                // 2) get total rod count for this hanger
                int rods = GetRodCount(fp);
                if (rods <= 0) continue;

                // 3) credit inserts:
                //    - if we have multiple diameter params, assign one insert to each in order
                //    - any remaining inserts (if rods > diameter entries) assume primary size
                int assigned = 0;
                for (int i = 0; i < diaStrings.Count && assigned < rods; ++i)
                {
                    Accumulate(perSize, diaStrings[i], 1);
                    assigned++;
                }
                if (assigned < rods)
                {
                    Accumulate(perSize, primary, rods - assigned);
                }

                grandTotal += rods;
            }

            // write CSV
            string project = GetProjectName(doc);
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string outDir = @"C:\Temp";
            try { if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir); } catch { }
            string fileName = $"{project}_{levelText}_Hanger Inserts_{date}.csv";
            string fullPath = Path.Combine(outDir, SafeFileName(fileName));

            try
            {
                WriteCsv(fullPath, project, levelText, perSize);
                TaskDialog.Show("ABMEP - Hanger Inserts",
                    $"CSV created:\n{fullPath}\n\nTotal Inserts (all sizes): {grandTotal}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ABMEP - Hanger Inserts", "Failed to write CSV:\n" + ex.Message);
                return Result.Failed;
            }
        }

        // ---------------- core counting helpers ----------------
        private static void Accumulate(Dictionary<string, int> map, string key, int add)
        {
            if (string.IsNullOrWhiteSpace(key)) key = "Unknown";
            if (!map.ContainsKey(key)) map[key] = 0;
            map[key] += Math.Max(0, add);
        }

        private static int GetRodCount(FabricationPart fp)
        {
            foreach (string n in RodCountParamNames)
            {
                var p = fp.LookupParameter(n);
                if (p == null || !p.HasValue) continue;

                try
                {
                    switch (p.StorageType)
                    {
                        case StorageType.Integer: return p.AsInteger();
                        case StorageType.Double: return (int)Math.Round(p.AsDouble());
                        case StorageType.String:
                            if (int.TryParse(p.AsString(), out int s)) return s;
                            break;
                    }
                }
                catch { }
            }
            return 1; // safe fallback
        }

        // ---------------- rod diameter logic (lifted from Trimble) ----------------
        private static List<string> GetRodDiameterStrings(FabricationPart fp)
        {
            var feet = GetRodDiametersFeetOrdered(fp);
            if (feet.Count == 0) feet = GetRodDiametersFeetFallback();

            var list = new List<string>(feet.Count);
            foreach (var f in feet)
            {
                double inches = f * 12.0;
                list.Add(InchesToNiceFraction(inches));
            }
            return list;
        }

        private static List<double> GetRodDiametersFeetOrdered(FabricationPart fp)
        {
            // read parameters like "rod_1_diameter", "rod 2 diameter", etc. (Double or String)
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
                        valFeet = p.AsDouble(); // Revit internal length (feet)
                    }
                    else if (p.StorageType == StorageType.String)
                    {
                        string s = p.AsString();
                        if (!TryParseFeetInches(s, out valFeet)) continue;
                    }

                    if (valFeet <= 0 || double.IsNaN(valFeet) || double.IsInfinity(valFeet)) continue;

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
                if (nameLower.Contains("rod_" + i) || nameLower.Contains("rod " + i) || nameLower.Contains("rod-" + i))
                    return i;
            return -1;
        }

        private static List<double> GetRodDiametersFeetFallback()
        {
            double In(double inches) => inches / 12.0;
            return new List<double> { In(0.375), In(0.5), In(0.625), In(0.75) };
        }

        private static string InchesToNiceFraction(double inches)
        {
            // Round to nearest 1/16"
            double frac = Math.Round(inches * 16.0) / 16.0;
            int whole = (int)Math.Floor(frac);
            double rem = frac - whole;

            int num = (int)Math.Round(rem * 16.0);
            int den = 16;
            if (num == 16) { whole += 1; num = 0; }

            if (num == 0) return whole.ToString(CultureInfo.InvariantCulture) + "\"";

            int g = GCD(num, den); num /= g; den /= g;
            return whole == 0
                ? $"{num}/{den}\""
                : $"{whole} {num}/{den}\"";
        }

        private static int GCD(int a, int b)
        {
            a = Math.Abs(a); b = Math.Abs(b);
            while (b != 0) { int t = a % b; a = b; b = t; }
            return a == 0 ? 1 : a;
        }

        // Parse strings like 0' 0 3/8" or 3/4" or 0.375
        private static bool TryParseFeetInches(string s, out double feet)
        {
            feet = 0;
            if (string.IsNullOrWhiteSpace(s)) return false;
            try
            {
                s = s.Replace("\"", "").Replace("in", "").Trim();
                double ft = 0, inch = 0;
                int idxFt = s.IndexOf("'");
                if (idxFt >= 0)
                {
                    double.TryParse(s.Substring(0, idxFt).Trim(), out ft);
                    s = s.Substring(idxFt + 1).Trim();
                }

                if (s.Contains("/"))
                {
                    var parts = s.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var part in parts)
                    {
                        if (part.Contains("/"))
                        {
                            var ab = part.Split('/');
                            if (ab.Length == 2 &&
                                double.TryParse(ab[0], out double a) &&
                                double.TryParse(ab[1], out double b) && b != 0)
                            {
                                inch += a / b;
                            }
                        }
                        else if (double.TryParse(part, out double val))
                        {
                            inch += val;
                        }
                    }
                }
                else
                {
                    // numeric inches (e.g. 0.375 or 3)
                    double.TryParse(s, out inch);
                }

                feet = ft + inch / 12.0;
                return feet > 0;
            }
            catch { return false; }
        }

        // ---------------- CSV helpers & UI ----------------
        private static void WriteCsv(string path, string project, string level, Dictionary<string, int> perSize)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Project,{Csv(project)}");
            sb.AppendLine($"Level,{Csv(level)}");
            sb.AppendLine();
            sb.AppendLine("Insert Size,Count");

            foreach (var kv in perSize.OrderBy(k => SizeSortKey(k.Key)))
                sb.AppendLine($"{Csv(kv.Key)},{kv.Value}");

            File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
        }

        private static double SizeSortKey(string label)
        {
            if (TryParseInchesLabel(label, out double inches)) return inches;
            return double.MaxValue;
        }

        private static bool TryParseInchesLabel(string label, out double inches)
        {
            inches = 0;
            if (string.IsNullOrWhiteSpace(label)) return false;
            string s = label.Trim().TrimEnd('"').Trim();

            if (double.TryParse(s, out double d)) { inches = d; return true; } // e.g. 0.375

            var parts = s.Split(' ');
            try
            {
                if (parts.Length == 2)
                {
                    int whole = int.Parse(parts[0]);
                    var frac = parts[1].Split('/');
                    inches = whole + (double.Parse(frac[0]) / double.Parse(frac[1]));
                    return true;
                }
                if (s.Contains("/"))
                {
                    var frac = s.Split('/');
                    inches = double.Parse(frac[0]) / double.Parse(frac[1]);
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static string Csv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static string SafeFileName(string name)
        {
            var bad = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name) sb.Append(Array.IndexOf(bad, ch) >= 0 ? '_' : ch);
            return sb.ToString().Trim();
        }

        private static string GetProjectName(Document doc)
        {
            try
            {
                var p = doc.ProjectInformation?.Name;
                if (!string.IsNullOrWhiteSpace(p)) return p.Trim();
            }
            catch { }
            return Path.GetFileNameWithoutExtension(doc.PathName);
        }

        private static string PromptLevel(string def)
        {
            try { return SimpleTextPrompt.Show("Hanger Inserts", "Enter Level text for the report:", def); }
            catch { return def; }
        }

        // tiny WinForms textbox prompt, like in Trimble
        private class SimpleTextPrompt : WinForms.Form
        {
            private readonly WinForms.TextBox _tb;
            private readonly WinForms.Button _ok, _cancel;

            private SimpleTextPrompt(string title, string message, string initial)
            {
                Text = title;
                StartPosition = WinForms.FormStartPosition.CenterScreen;
                FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                MinimizeBox = false; MaximizeBox = false;
                Width = 420; Height = 160;

                var lbl = new WinForms.Label { Left = 12, Top = 12, Width = 380, Text = message };
                _tb = new WinForms.TextBox { Left = 12, Top = 38, Width = 380, Text = initial ?? "" };
                _ok = new WinForms.Button { Text = "OK", Left = 226, Width = 80, Top = 70, DialogResult = WinForms.DialogResult.OK };
                _cancel = new WinForms.Button { Text = "Cancel", Left = 312, Width = 80, Top = 70, DialogResult = WinForms.DialogResult.Cancel };

                Controls.Add(lbl); Controls.Add(_tb); Controls.Add(_ok); Controls.Add(_cancel);
                AcceptButton = _ok; CancelButton = _cancel;
            }

            public static string Show(string title, string message, string initial)
            {
                using (var f = new SimpleTextPrompt(title, message, initial))
                {
                    var r = f.ShowDialog();
                    return r == WinForms.DialogResult.OK ? f._tb.Text : initial;
                }
            }
        }
    }
}
