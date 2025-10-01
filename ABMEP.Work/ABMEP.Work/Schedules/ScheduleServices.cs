// File: ABMEP.Work/Schedules/ScheduleServices.cs
// Target: .NET Framework 4.8
// Namespace: ABMEP.Work.Schedules
//
// What this does (no other files required):
// - Creates a true MEP Fabrication schedule (Pipework or Ductwork) for a given assembly
// - Copies fields/formatting from a best prototype or from a named prototype if supplied
// - Filters the schedule to the assembly by name (robust field-name discovery)
// - Places the schedule either to the "Spool Window" annotation's top-left or, if not found,
//   1/16" from the left and 1/8" from the top of the title block.

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Work.Schedules
{
    internal static class Units
    {
        public const double INCH = 1.0 / 12.0;
    }

    public static class ScheduleServices
    {
        // Entry: create + configure + place
        public static ViewSchedule CreateFabricationScheduleForAssembly(
            Document doc,
            AssemblyInstance asm,
            string prototypeScheduleNameOrNull,
            string viewTemplateNameOrNull)
        {
            if (doc == null || asm == null) return null;

            // 1) Choose category based on members
            ElementId catId = FabricationCategory.DetectForAssembly(doc, asm);

            // 2) New schedule
            var sch = ViewSchedule.CreateSchedule(doc, catId);
            var asmName = GetAssemblyName(doc, asm);
            sch.Name = $"{asmName} – Schedule";

            // 3) Apply template (formatting)
            var vt = ScheduleFinder.FindTemplateByExactName(doc, viewTemplateNameOrNull);
            if (vt != null) sch.ViewTemplateId = vt.Id;

            // 4) Copy fields from best prototype (to match your chosen schedule look)
            var proto = ScheduleFinder.BestPrototypeFor(doc, sch, prototypeScheduleNameOrNull);
            if (proto != null) ScheduleFieldSync.CopyFieldsFrom(doc, sch, proto);

            // 5) Filter to the assembly
            AssemblyFilter.ApplyByAssemblyName(sch, doc, asmName);

            return sch;
        }

        public static ScheduleSheetInstance PlaceScheduleOnSheet(
            Document doc,
            ViewSheet sheet,
            ViewSchedule schedule)
        {
            if (doc == null || sheet == null || schedule == null) return null;

            // Try Spool Window first
            var wnd = SpoolWindowFinder.FindOnSheet(doc, sheet);
            if (wnd != null)
            {
                // Top-left of the annotation bbox (on sheets: Min.X, Max.Y)
                var bb = wnd.get_BoundingBox(sheet);
                if (bb != null)
                {
                    var targetTL = new XYZ(bb.Min.X, bb.Max.Y, 0.0);
                    var inst = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, XYZ.Zero);
                    var sbb = inst.get_BoundingBox(sheet);
                    if (sbb != null)
                    {
                        var curTL = new XYZ(sbb.Min.X, sbb.Max.Y, 0.0);
                        ElementTransformUtils.MoveElement(doc, inst.Id, targetTL - curTL);
                    }
                    return inst;
                }
            }

            // Fallback: place relative to title block with 1/16" left & 1/8" top
            return SchedulePlacer.Place(doc, sheet, schedule, 1.0 / 16.0, 1.0 / 8.0);
        }

        // Helper to read assembly display name
        private static string GetAssemblyName(Document doc, AssemblyInstance asm)
        {
            var et = doc.GetElement(asm.GetTypeId()) as ElementType;
            if (et != null && !string.IsNullOrWhiteSpace(et.Name)) return et.Name;
            return asm.Name ?? asm.Id.IntegerValue.ToString();
        }
    }

    // ---------------- internals ----------------

    internal static class FabricationCategory
    {
        public static ElementId DetectForAssembly(Document doc, AssemblyInstance asm)
        {
            if (doc == null || asm == null) return new ElementId(BuiltInCategory.OST_FabricationPipework);

            int pipeHits = 0, ductHits = 0;
            foreach (var mid in asm.GetMemberIds())
            {
                var el = doc.GetElement(mid);
                var cat = el?.Category?.Id;
                if (cat == null) continue;
                if (cat.IntegerValue == (int)BuiltInCategory.OST_FabricationPipework) pipeHits++;
                if (cat.IntegerValue == (int)BuiltInCategory.OST_FabricationDuctwork) ductHits++;
            }
            return (pipeHits >= ductHits)
                ? new ElementId(BuiltInCategory.OST_FabricationPipework)
                : new ElementId(BuiltInCategory.OST_FabricationDuctwork);
        }
    }

    internal static class ScheduleFinder
    {
        public static View FindTemplateByExactName(Document doc, string name)
        {
            if (doc == null || string.IsNullOrWhiteSpace(name)) return null;
            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .FirstOrDefault(v => v.IsTemplate && v.ViewType == ViewType.Schedule &&
                                     string.Equals(v.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        public static ViewSchedule FindPrototypeByExactName(Document doc, string name, ElementId categoryId)
        {
            if (doc == null || string.IsNullOrWhiteSpace(name)) return null;
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate &&
                             vs.Definition != null &&
                             vs.Definition.CategoryId != null &&
                             vs.Definition.CategoryId.IntegerValue == categoryId.IntegerValue)
                .FirstOrDefault(vs => string.Equals(vs.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        public static IEnumerable<ViewSchedule> AllPrototypesInCategory(Document doc, ElementId categoryId)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate &&
                             vs.Definition != null &&
                             vs.Definition.CategoryId != null &&
                             vs.Definition.CategoryId.IntegerValue == categoryId.IntegerValue);
        }

        public static ViewSchedule BestPrototypeFor(Document doc, ViewSchedule target, string preferredName)
        {
            if (doc == null || target?.Definition == null) return null;
            var catId = target.Definition.CategoryId ?? new ElementId(BuiltInCategory.OST_FabricationPipework);

            var exact = FindPrototypeByExactName(doc, preferredName, catId);
            if (exact != null) return exact;

            ViewSchedule best = null;
            int bestCount = -1;
            foreach (var vs in AllPrototypesInCategory(doc, catId))
            {
                int c = vs.Definition?.GetFieldCount() ?? 0;
                if (c > bestCount) { best = vs; bestCount = c; }
            }
            return best;
        }
    }

    internal static class ScheduleFieldSync
    {
        public static void CopyFieldsFrom(Document doc, ViewSchedule target, ViewSchedule prototype)
        {
            if (doc == null || target == null || prototype == null) return;

            var tdef = target.Definition;
            var pdef = prototype.Definition;
            if (tdef == null || pdef == null) return;

            // remove whatever Revit added
            var remove = new List<ScheduleFieldId>();
            for (int i = 0; i < tdef.GetFieldCount(); i++)
                remove.Add(tdef.GetField(i).FieldId);
            for (int i = remove.Count - 1; i >= 0; i--)
                tdef.RemoveField(remove[i]);

            IList<SchedulableField> targetFields = tdef.GetSchedulableFields();

            string NameOf(SchedulableField sf) => sf?.GetName(doc) ?? "";

            // replicate prototype order/headings
            for (int i = 0; i < pdef.GetFieldCount(); i++)
            {
                var pf = pdef.GetField(i);

                SchedulableField match = null;

                // strong match by ParameterId
                if (pf.ParameterId != null && pf.ParameterId.IntegerValue != -1)
                {
                    match = targetFields.FirstOrDefault(sf =>
                        sf.ParameterId != null &&
                        sf.ParameterId.IntegerValue == pf.ParameterId.IntegerValue);
                }

                // fallback match by name
                if (match == null)
                {
                    var want = pf.GetName();
                    match = targetFields.FirstOrDefault(sf =>
                        string.Equals(NameOf(sf), want, StringComparison.OrdinalIgnoreCase));
                }

                if (match == null) continue;

                var tf = tdef.AddField(match);
                try { tf.ColumnHeading = pf.ColumnHeading; } catch { }
                try { tf.IsHidden = pf.IsHidden; } catch { }
            }
        }
    }

    internal static class AssemblyFilter
    {
        // Common field-name patterns used by different shops/content
        private static readonly string[] Candidates =
        {
            "Item Spool",
            "Spool Name",
            "Spool",
            "Assembly Name",
            "Assembly"
        };

        public static void ApplyByAssemblyName(ViewSchedule schedule, Document doc, string assemblyName)
        {
            if (schedule?.Definition == null || doc == null || string.IsNullOrWhiteSpace(assemblyName)) return;

            var def = schedule.Definition;
            var sfs = def.GetSchedulableFields();

            // Find a schedulable field whose display name contains one of our candidates
            SchedulableField chosen = null;
            foreach (var sf in sfs)
            {
                string n = sf.GetName(doc) ?? "";
                if (Candidates.Any(c => n.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0))
                { chosen = sf; break; }
            }

            if (chosen == null)
            {
                // Instrumentation: show what *is* available so we can lock the right one next pass
                var available = string.Join("\n", sfs.Select(f => " • " + (f.GetName(doc) ?? "<unnamed>")));
                TaskDialog.Show("ABMEP – Schedule Filter",
                    "I couldn’t find a field to filter by the assembly/spool name.\n\n" +
                    "Available schedulable field names in your Fabrication schedule:\n\n" +
                    available +
                    "\n\nTell me which one corresponds to the spool/assembly, and I’ll wire it in.");
                return;
            }

            // Ensure the field is part of the schedule (hidden is fine)
            ScheduleField sfInSchedule = null;
            for (int i = 0; i < def.GetFieldCount(); i++)
            {
                var f = def.GetField(i);
                if (f.ParameterId != null &&
                    chosen.ParameterId != null &&
                    f.ParameterId.IntegerValue == chosen.ParameterId.IntegerValue)
                { sfInSchedule = f; break; }
            }
            if (sfInSchedule == null) sfInSchedule = def.AddField(chosen);
            try { sfInSchedule.IsHidden = true; } catch { }

            def.ClearFilters();
            def.AddFilter(new ScheduleFilter(sfInSchedule.FieldId, ScheduleFilterType.Equal, assemblyName));
        }
    }

    internal static class SpoolWindowFinder
    {
        // Looks for a Generic Annotation instance named "Spool Window" on the sheet
        public static FamilyInstance FindOnSheet(Document doc, ViewSheet sheet)
        {
            try
            {
                return new FilteredElementCollector(doc, sheet.Id)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .FirstOrDefault(fi =>
                    {
                        var cat = fi.Category;
                        if (cat == null || cat.Id.IntegerValue != (int)BuiltInCategory.OST_GenericAnnotation) return false;
                        var sym = doc.GetElement(fi.GetTypeId()) as ElementType;
                        var famName = sym?.FamilyName ?? "";
                        var typeName = sym?.Name ?? "";
                        // match either family or type naming to "Spool Window"
                        return famName.IndexOf("Spool Window", StringComparison.OrdinalIgnoreCase) >= 0
                            || typeName.IndexOf("Spool Window", StringComparison.OrdinalIgnoreCase) >= 0;
                    });
            }
            catch { return null; }
        }
    }

    internal static class SchedulePlacer
    {
        public static ScheduleSheetInstance Place(Document doc, ViewSheet sheet, ViewSchedule schedule,
                                                  double leftInches, double topInches)
        {
            if (doc == null || sheet == null || schedule == null) return null;

            var tb = new FilteredElementCollector(doc, sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .FirstOrDefault();
            if (tb == null) return null;

            var bb = tb.get_BoundingBox(sheet);
            if (bb == null) return null;

            double left = bb.Min.X;
            double top = bb.Max.Y;
            var targetTL = new XYZ(left + leftInches * Units.INCH, top - topInches * Units.INCH, 0);

            var inst = ScheduleSheetInstance.Create(doc, sheet.Id, schedule.Id, XYZ.Zero);
            var sbb = inst.get_BoundingBox(sheet);
            if (sbb != null)
            {
                var curTL = new XYZ(sbb.Min.X, sbb.Max.Y, 0.0);
                ElementTransformUtils.MoveElement(doc, inst.Id, targetTL - curTL);
            }
            return inst;
        }
    }
}
