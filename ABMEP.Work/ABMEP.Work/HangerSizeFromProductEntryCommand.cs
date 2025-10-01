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
    public class HangerSizeFromProductEntryCommand : IExternalCommand
    {
        private const string SOURCE_PARAM = "Product Entry";
        private const string TARGET_PARAM = "Hanger Size";

        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            UIDocument uidoc = data.Application.ActiveUIDocument;
            Document doc = uidoc?.Document;
            if (doc == null)
            {
                message = "No active document.";
                return Result.Failed;
            }

            // Use current selection; if empty, let user pick
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds(); // <-- ICollection (not ISet)
            if (ids == null || ids.Count == 0)
            {
                try
                {
                    var refs = uidoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Select items to copy 'Product Entry' → 'Hanger Size'. Press ESC when done.");

                    // List<ElementId> implements ICollection<ElementId>
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
                TaskDialog.Show("Hanger Size", "No elements selected.");
                return Result.Cancelled;
            }

            int updated = 0, skippedNoSource = 0, skippedNoTarget = 0, skippedReadonly = 0;
            var errors = new List<string>();

            using (var t = new Transaction(doc, "Copy Product Entry → Hanger Size"))
            {
                t.Start();

                foreach (var e in elems)
                {
                    try
                    {
                        string src = GetParamString(e, SOURCE_PARAM);
                        if (string.IsNullOrWhiteSpace(src)) { skippedNoSource++; continue; }

                        Parameter target = e.LookupParameter(TARGET_PARAM);
                        if (target == null) { skippedNoTarget++; continue; }
                        if (target.IsReadOnly) { skippedReadonly++; continue; }

                        string val = src.Trim();
                        bool ok = false;

                        if (target.StorageType == StorageType.String)
                        {
                            ok = target.Set(val);
                        }
                        else
                        {
                            try { ok = target.SetValueString(val); }
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
            if (skippedNoSource > 0) sb.AppendLine($"Skipped (no '{SOURCE_PARAM}'): {skippedNoSource}");
            if (skippedNoTarget > 0) sb.AppendLine($"Skipped (no '{TARGET_PARAM}'): {skippedNoTarget}");
            if (skippedReadonly > 0) sb.AppendLine($"Skipped (read-only '{TARGET_PARAM}'): {skippedReadonly}");
            if (errors.Count > 0)
            {
                sb.AppendLine().AppendLine("Errors:");
                foreach (var s in errors.Take(10)) sb.AppendLine(" • " + s);
                if (errors.Count > 10) sb.AppendLine($" • +{errors.Count - 10} more…");
            }

            TaskDialog.Show("Hanger Size", sb.ToString());
            return Result.Succeeded;
        }

        // --- helpers ---

        private static string GetParamString(Element e, string name)
        {
            Parameter p = e.LookupParameter(name);
            if (p == null) return null;

            try
            {
                if (p.StorageType == StorageType.String)
                    return p.AsString();

                // Prefer formatted value if available
                string vs = null;
                try { vs = p.AsValueString(); } catch { }
                if (!string.IsNullOrWhiteSpace(vs)) return vs;

                switch (p.StorageType)
                {
                    case StorageType.Integer: return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                    case StorageType.Double: return p.AsDouble().ToString("0.########", CultureInfo.InvariantCulture);
                    case StorageType.ElementId: return p.AsElementId().Value.ToString(CultureInfo.InvariantCulture);
                    default: return null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static string ElementDesc(Element e)
        {
            try
            {
                string n = e.Name ?? "";
                string cat = e.Category?.Name ?? "";
                return $"[{cat}] {n} (Id {e.Id.Value})";
            }
            catch
            {
                return $"(Id {e?.Id.Value})";
            }
        }
    }
}
