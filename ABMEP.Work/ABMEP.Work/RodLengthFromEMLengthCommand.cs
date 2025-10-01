// Target: .NET Framework 4.8
// Assembly: ABMEP.Work.dll

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class RodLengthFromEMLengthsCommand : IExternalCommand
    {
        private const string SRC_A = "eM_Length A";
        private const string SRC_B = "eM_Length B";
        private const string TARGET = "Rod Length";

        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            UIDocument uidoc = data.Application.ActiveUIDocument;
            Document doc = uidoc?.Document;
            if (doc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }

            // Current selection; if empty, prompt
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids == null || ids.Count == 0)
            {
                try
                {
                    var picks = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Select hangers/items to sum 'eM_Length A' + 'eM_Length B' → 'Rod Length'. Press ESC when done.");
                    ids = picks.Select(r => r.ElementId).ToList();
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }
            }

            var elems = ids.Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
            if (elems.Count == 0)
            {
                TaskDialog.Show("Rod Length", "No elements selected.");
                return Result.Cancelled;
            }

            int updated = 0, skippedNoSource = 0, skippedNoTarget = 0, skippedReadonly = 0;
            var errors = new List<string>();

            using (var t = new Transaction(doc, "eM_Length A+B → Rod Length"))
            {
                t.Start();

                foreach (var e in elems)
                {
                    try
                    {
                        // Read both source values
                        var haveA = TryReadParamFeetOrText(e.LookupParameter(SRC_A), out double aFt, out string aTxt);
                        var haveB = TryReadParamFeetOrText(e.LookupParameter(SRC_B), out double bFt, out string bTxt);

                        if (!haveA && !haveB)
                        {
                            skippedNoSource++;
                            continue;
                        }

                        // Build a summed value. Prefer numeric feet when available; otherwise parse text.
                        double totalFt = 0;
                        bool anyFeet = false;

                        if (haveA && IsFinite(aFt)) { totalFt += aFt; anyFeet = true; }
                        else if (haveA && TryParseFeetInches(aTxt, out double aParsed)) { totalFt += aParsed; anyFeet = true; }

                        if (haveB && IsFinite(bFt)) { totalFt += bFt; anyFeet = true; }
                        else if (haveB && TryParseFeetInches(bTxt, out double bParsed)) { totalFt += bParsed; anyFeet = true; }

                        // If we still don't have a numeric result, try concatenating formatted texts then parsing (edge case).
                        if (!anyFeet)
                        {
                            string merged = $"{aTxt} + {bTxt}";
                            if (!TryParseFeetInches(aTxt, out double af) && !TryParseFeetInches(bTxt, out double bf))
                            {
                                // We truly have no usable numeric data
                                skippedNoSource++;
                                continue;
                            }
                        }

                        Parameter dst = e.LookupParameter(TARGET);
                        if (dst == null) { skippedNoTarget++; continue; }
                        if (dst.IsReadOnly) { skippedReadonly++; continue; }

                        bool ok = false;

                        if (dst.StorageType == StorageType.Double)
                        {
                            ok = dst.Set(totalFt); // internal feet
                        }
                        else if (dst.StorageType == StorageType.String)
                        {
                            ok = dst.Set(FeetToFeetInches(totalFt));
                        }
                        else
                        {
                            try { ok = dst.SetValueString(FeetToFeetInches(totalFt)); }
                            catch { ok = false; }
                        }

                        if (ok) updated++;
                        else errors.Add($"Could not set '{TARGET}' on {ElementDesc(e)}.");
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"{ElementDesc(e)}: {ex.Message}");
                    }
                }

                t.Commit();
            }

            var sb = new StringBuilder();
            sb.AppendLine($"Updated: {updated}");
            if (skippedNoSource > 0) sb.AppendLine($"Skipped (no '{SRC_A}' & '{SRC_B}'): {skippedNoSource}");
            if (skippedNoTarget > 0) sb.AppendLine($"Skipped (no '{TARGET}'): {skippedNoTarget}");
            if (skippedReadonly > 0) sb.AppendLine($"Skipped (read-only '{TARGET}'): {skippedReadonly}");
            if (errors.Count > 0)
            {
                sb.AppendLine().AppendLine("Errors:");
                foreach (var s in errors.Take(10)) sb.AppendLine(" • " + s);
                if (errors.Count > 10) sb.AppendLine($" • +{errors.Count - 10} more…");
            }

            TaskDialog.Show("ABMEP - Rod Length", sb.ToString());
            return Result.Succeeded;
        }

        // -------- parameter readers / format helpers --------

        private static bool TryReadParamFeetOrText(Parameter p, out double feet, out string text)
        {
            feet = double.NaN;
            text = null;
            if (p == null || !p.HasValue) return false;

            if (p.StorageType == StorageType.Double)
            {
                feet = p.AsDouble(); // internal feet
                try { text = p.AsValueString(); } catch { text = null; }
                return true;
            }

            // Prefer formatted value; fall back to string
            try { text = p.AsValueString(); } catch { text = null; }
            if (string.IsNullOrWhiteSpace(text) && p.StorageType == StorageType.String)
                text = p.AsString();

            if (!string.IsNullOrWhiteSpace(text) && TryParseFeetInches(text, out double ft))
                feet = ft;

            return !string.IsNullOrWhiteSpace(text) || (IsFinite(feet) && feet > 0);
        }

        private static bool IsFinite(double v) => !(double.IsNaN(v) || double.IsInfinity(v));

        private static string ElementDesc(Element e)
        {
            try
            {
                string n = e.Name ?? "";
                string cat = e.Category?.Name ?? "";
                return $"[{cat}] {n} (Id {e.Id.Value})";
            }
            catch { return $"(Id {e?.Id.Value})"; }
        }

        // Format feet as 2' - 0 3/8"
        private static string FeetToFeetInches(double feet)
        {
            if (!IsFinite(feet)) return "";
            double inchesTotal = feet * 12.0;
            int feetPart = (int)Math.Floor(inchesTotal / 12.0);
            double inchRema = inchesTotal - feetPart * 12.0;

            int sixteenths = (int)Math.Round(inchRema * 16.0);
            int inchWhole = sixteenths / 16;
            int fracNum = sixteenths % 16;
            int fracDen = 16;

            if (fracNum != 0)
            {
                int g = GCD(fracNum, fracDen);
                fracNum /= g; fracDen /= g;
            }

            if (inchWhole >= 12) { feetPart += 1; inchWhole -= 12; }

            var sb = new StringBuilder();
            sb.Append(feetPart.ToString(CultureInfo.InvariantCulture)).Append("' - ");
            if (fracNum == 0)
                sb.Append(inchWhole.ToString(CultureInfo.InvariantCulture)).Append("\"");
            else
                sb.Append(inchWhole.ToString(CultureInfo.InvariantCulture)).Append(" ")
                  .Append(fracNum.ToString(CultureInfo.InvariantCulture)).Append("/")
                  .Append(fracDen.ToString(CultureInfo.InvariantCulture)).Append("\"");
            return sb.ToString();
        }

        private static int GCD(int a, int b)
        {
            a = Math.Abs(a); b = Math.Abs(b);
            while (b != 0) { int t = a % b; a = b; b = t; }
            return a == 0 ? 1 : a;
        }

        // Robust best-effort parser for 2' - 0 3/8", 24", 2 ft, etc.
        private static bool TryParseFeetInches(string s, out double feet)
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

                // crude mm detection
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
    }
}
