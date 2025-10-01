// ABMEP.Work / Services / DimensionsToSleevesService.cs
// Target: .NET Framework 4.8 | Revit 2024+
//
// Update:
// - Include BOTH Structural Columns (OST_StructuralColumns) and Architectural Columns (OST_Columns).
// - If no columns are found OR column-grid filter produces zero, fall back to ALL grids
//   so the command never returns "No qualifying sleeves / grids found" just because
//   columns are in a link or missing in host.
// - Keeps the “only dimension to column grids” behavior when host columns exist.

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace ABMEP.Work.Services
{
    public sealed class DimensionsToSleevesService
    {
        private readonly Document _doc;

        // Slight offset so dimension line is not on top of sleeve graphics (feet)
        private const double DIM_OFFSET_FT = 0.15;

        public DimensionsToSleevesService(Document doc) => _doc = doc;

        // Back-compat overloads
        public DimensionsToSleevesService(Document doc, View _ /*ignored*/) => _doc = doc;
        public int Run(View _ /*ignored*/) => Run();

        public int Run()
        {
            var sleeves = GetSleeves();
            var allGrids = GetAllGrids();
            var hostColumns = GetAllHostColumns();

            if (sleeves.Count == 0 || allGrids.Count == 0)
                return 0;

            // Prefer grids that intersect host columns; if none found, fall back to all grids.
            var qualifying = (hostColumns.Count > 0)
                ? FilterGridsThatHitAnyColumn(allGrids, hostColumns, _doc.Application.ShortCurveTolerance)
                : new List<Grid>();

            if (qualifying.Count == 0)
                qualifying = allGrids; // fallback: use all grids so we still place dims

            var (verticalGrids, horizontalGrids) = SplitGridsByOrientation(qualifying);

            var dimType = FindLinearDimTypeByName("1/4 Lee Dimension Linear");
            if (dimType == null) return 0;

            int placed = 0;

            using (var tx = new Transaction(_doc, "ABMEP – Dimension Sleeves to Grids"))
            {
                tx.Start();

                foreach (var sleeve in sleeves)
                {
                    var lp = sleeve.Location as LocationPoint;
                    if (lp == null) continue;
                    XYZ p = lp.Point;

                    var refLR = TryGetReference(sleeve, FamilyInstanceReferenceType.CenterLeftRight);
                    var refFB = TryGetReference(sleeve, FamilyInstanceReferenceType.CenterFrontBack);

                    Grid nearestV = NearestGridToPoint(verticalGrids, p);
                    if (nearestV != null && refLR != null)
                    {
                        var dimLine = BuildInfiniteLineThrough(p + DIM_OFFSET_FT * XYZ.BasisY, XYZ.BasisX);
                        if (TryMakeDim(nearestV, refLR, dimLine, dimType)) placed++;
                    }

                    Grid nearestH = NearestGridToPoint(horizontalGrids, p);
                    if (nearestH != null && refFB != null)
                    {
                        var dimLine = BuildInfiniteLineThrough(p + DIM_OFFSET_FT * XYZ.BasisX, XYZ.BasisY);
                        if (TryMakeDim(nearestH, refFB, dimLine, dimType)) placed++;
                    }
                }

                tx.Commit();
            }

            return placed;
        }

        // ---------- collectors ----------

        private List<FamilyInstance> GetSleeves()
        {
            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_ConduitFitting)
                .WhereElementIsNotElementType()
                .OfType<FamilyInstance>()
                .ToList();
        }

        private List<Grid> GetAllGrids()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .Where(g => g.Curve != null)
                .ToList();
        }

        /// <summary>Columns that live in THIS document (both structural & architectural).</summary>
        private List<FamilyInstance> GetAllHostColumns()
        {
            var list = new List<FamilyInstance>();

            // Structural Columns
            list.AddRange(new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .OfType<FamilyInstance>());

            // Architectural Columns (Note: category is OST_Columns)
            list.AddRange(new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Columns)
                .WhereElementIsNotElementType()
                .OfType<FamilyInstance>());

            return list;
        }

        // ---------- qualifying grids ----------

        private static List<Grid> FilterGridsThatHitAnyColumn(
            IEnumerable<Grid> grids,
            IEnumerable<FamilyInstance> columns,
            double tol)
        {
            var result = new List<Grid>();
            foreach (var g in grids)
            {
                var c = g.Curve;
                if (c == null) continue;

                bool hit = false;
                foreach (var col in columns)
                {
                    if (col.Location is LocationPoint lp && lp.Point != null)
                    {
                        var proj = c.Project(lp.Point);
                        if (proj != null && proj.XYZPoint != null &&
                            proj.XYZPoint.DistanceTo(lp.Point) <= tol)
                        {
                            hit = true;
                            break;
                        }
                    }
                }
                if (hit) result.Add(g);
            }
            return result;
        }

        private static (List<Grid> vertical, List<Grid> horizontal) SplitGridsByOrientation(IEnumerable<Grid> grids)
        {
            var vertical = new List<Grid>();
            var horizontal = new List<Grid>();

            foreach (var g in grids)
            {
                var curve = g.Curve;
                if (curve == null) continue;

                XYZ d = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();
                // Treat “more Y than X” as vertical (north–south)
                if (Math.Abs(d.Y) >= Math.Abs(d.X)) vertical.Add(g); else horizontal.Add(g);
            }
            return (vertical, horizontal);
        }

        // ---------- dims helpers ----------

        private DimensionType FindLinearDimTypeByName(string name)
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(DimensionType))
                .Cast<DimensionType>()
                .FirstOrDefault(dt =>
                    dt != null &&
                    dt.StyleType == DimensionStyleType.Linear &&
                    string.Equals(dt.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        private static Reference TryGetReference(FamilyInstance fi, FamilyInstanceReferenceType rtype)
        {
            try
            {
                var refs = fi.GetReferences(rtype);
                return refs != null ? refs.FirstOrDefault() : null;
            }
            catch { return null; }
        }

        private static Grid NearestGridToPoint(IEnumerable<Grid> grids, XYZ p)
        {
            double best = double.MaxValue;
            Grid winner = null;

            foreach (var g in grids)
            {
                var proj = g.Curve?.Project(p);
                if (proj == null) continue;
                double d = proj.XYZPoint.DistanceTo(p);
                if (d < best) { best = d; winner = g; }
            }
            return winner;
        }

        private static Line BuildInfiniteLineThrough(XYZ origin, XYZ dir)
        {
            var u = dir.Normalize();
            double L = 1000.0; // feet
            return Line.CreateBound(origin - L * u, origin + L * u);
        }

        private bool TryMakeDim(Grid grid, Reference sleeveRef, Line dimLine, DimensionType dimType)
        {
            try
            {
                var rarr = new ReferenceArray();
                rarr.Append(sleeveRef);
                rarr.Append(new Reference(grid));
                var dim = _doc.Create.NewDimension(_doc.ActiveView, dimLine, rarr, dimType);
                return dim != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
