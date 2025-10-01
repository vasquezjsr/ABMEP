// Target: .NET Framework 4.8
// Assembly: ABMEP.Work.dll

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Fabrication;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class StrutLengthFromWidthCommand : IExternalCommand
    {
        private const string SOURCE_PARAM_PRIMARY = "eM_Width";   // eVolve Mechanical
        private const string SOURCE_PARAM_FALLBACK = "Width";     // (optional) fallback
        private const string TARGET_PARAM = "Strut Length";

        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            UIDocument uidoc = data.Application.ActiveUIDocument;
            Document doc = uidoc?.Document;
            if (doc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }

            // Use current selection; if empty, prompt
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            if (ids == null || ids.Count == 0)
            {
                try
                {
                    var refs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Select items to copy 'eM_Width' → 'Strut Length'. Press ESC when done.");
                    ids = refs.Select(r => r.ElementId).ToList();
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    return Result.Cancelled;
                }
            }

            var elems = ids.Select(id => doc.GetElement(id)).Where(e => e != null).ToList();
            if (elems.Count == 0)
            {
                TaskDialog.Show("Strut Length", "No elements selected.");
                return Result.Cancelled;
            }

            int updated = 0, skippedNoSource = 0, skippedNoTarget = 0, skippedReadonly = 0;
            var errors = new List<string>();

            using (var t = new Transaction(doc, "Copy eM_Width → Strut Length"))
            {
                t.Start();

                foreach (var e in elems)
                {
                    try
                    {
                        // Get source (prefer eM_Width)
                        if (!TryGetWidthFeetOrText(e, out double widthFt, out string widthText))
                        {
                            skippedNoSource++;
                            continue;
                        }

                        Parameter dst = e.LookupParameter(TARGET_PARAM);
                        if (dst == null) { skippedNoTarget++; continue; }
                        if (dst.IsReadOnly) { skippedReadonly++; continue; }

                        bool ok = false;

                        // If destination is a Length/double param, push raw feet when we have it
                        if (dst.StorageType == StorageType.Double && IsFinite(widthFt) && widthFt > 0)
                        {
                            ok = dst.Set(widthFt);
                        }
                        else if (dst.StorageType == StorageType.String)
                        {
                            string s = IsFinite(widthFt) && widthFt > 0
                                ? FeetToFeetInches(widthFt)
                                : (widthText ?? "");
                            ok = dst.Set(s ?? "");
                        }
                        else
                        {
                            string s = IsFinite(widthFt) && widthFt > 0
                                ? FeetToFeetInches(widthFt)
                                : (widthText ?? "");
                            try { ok = dst.SetValueString(s ?? ""); }
                            catch { ok = false; }
                        }

                        if (ok) updated++;
                        else errors.Add($"Could not set '{TARGET_PARAM}' on {ElementDesc(e)}.");
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
            if (skippedNoSource > 0) sb.AppendLine($"Skipped (no 'eM_Width'): {skippedNoSource}");
            if (skippedNoTarget > 0) sb.AppendLine($"Skipped (no '{TARGET_PARAM}'): {skippedNoTarget}");
            if (skippedReadonly > 0) sb.AppendLine($"Skipped (read-only '{TARGET_PARAM}'): {skippedReadonly}");
            if (errors.Count > 0)
            {
                sb.AppendLine().AppendLine("Errors:");
                foreach (var s in errors.Take(10)) sb.AppendLine(" • " + s);
                if (errors.Count > 10) sb.AppendLine($" • +{errors.Count - 10} more…");
            }

            TaskDialog.Show("ABMEP - Strut Length", sb.ToString());
            return Result.Succeeded;
        }

        // -------- width fetch --------
        private static bool TryGetWidthFeetOrText(Element e, out double widthFeet, out string widthText)
        {
            widthFeet = double.NaN;
            widthText = null;

            // 1) eVolve Mechanical: eM_Width
            if (TryReadParamFeetOrText(e.LookupParameter(SOURCE_PARAM_PRIMARY), out widthFeet, out widthText))
                return true;

            // 2) Optional fallback to a plain "Width" parameter
            if (TryReadParamFeetOrText(e.LookupParameter(SOURCE_PARAM_FALLBACK), out widthFeet, out widthText))
                return true;

            // 3) Last resort: fabrication dimension named "Width" (for some patterns)
            if (e is FabricationPart fp)
            {
                try
                {
                    var dimsMethod = typeof(FabricationPart).GetMethod("GetDimensions");
                    if (dimsMethod != null)
                    {
                        var dimsEnum = dimsMethod.Invoke(fp, null) as System.Collections.IEnumerable;
                        foreach (var dim in dimsEnum)
                        {
                            var t = dim.GetType();
                            string name = t.GetProperty("Name")?.GetValue(dim) as string;
                            if (!string.Equals(name, "Width", StringComparison.OrdinalIgnoreCase)) continue;

                            object val = t.GetProperty("Value")?.GetValue(dim);
                            if (val is double dv) { widthFeet = dv; widthText = FeetToFeetInches(dv); return true; }
                            if (val != null && double.TryParse(val.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double dv2))
                            { widthFeet = dv2; widthText = FeetToFeetInches(dv2); return true; }

                            object ftxt = t.GetProperty("DisplayValue")?.GetValue(dim);
                            if (ftxt is string sTxt)
                            {
                                widthText = sTxt;
                                if (TryParseFeetInches(sTxt, out double fv)) widthFeet = fv;
                                return true;
                            }
                        }
                    }
                }
                catch { /* ignore */ }
            }

            return false;
        }

        private static bool TryReadParamFeetOrText(Parameter p, out double feet, out string text)
        {
            feet = double.NaN;
            text = null;
            if (p == null || !p.HasValue) return false;

            if (p.StorageType == StorageType.Double)
            {
                feet = p.AsDouble();            // internal feet
                try { text = p.AsValueString(); } catch { text = null; }
                return true;
            }

            // Try formatted value first
            try { text = p.AsValueString(); } catch { text = null; }
            if (string.IsNullOrWhiteSpace(text) && p.StorageType == StorageType.String)
                text = p.AsString();

            if (!string.IsNullOrWhiteSpace(text) && TryParseFeetInches(text, out double ft))
                feet = ft;

            return !string.IsNullOrWhiteSpace(text) || (IsFinite(feet) && feet > 0);
        }

        // -------- utilities --------
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

        // Robust parser for 2' - 0 3/8", 24", 2 ft, etc. (best effort)
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

                // crude mm check
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
