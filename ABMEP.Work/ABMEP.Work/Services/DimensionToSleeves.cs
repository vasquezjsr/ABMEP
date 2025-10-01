// Target: .NET Framework 4.8
// Revit: 2024+
// Namespace: ABMEP.Work.Commands

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ABMEP.Work.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class DimensionsToSleeves : IExternalCommand
    {
        private const string DialogTitle = "ABMEP";

        private const double Tolerance = 1e-6;
        private const double VeryLarge = 1e+9;
        private const double MinFaceArea = 1.0 / 144.0; // 1 in² in ft²

        public Result Execute(ExternalCommandData c, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = c.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                if (uidoc == null || uidoc.Document == null)
                {
                    message = "No active document.";
                    return Result.Failed;
                }

                Document doc = uidoc.Document;
                View view = uidoc.ActiveView;
                if (!(view is ViewPlan))
                {
                    TaskDialog.Show(DialogTitle, "Please run Dimensions to Sleeves in a plan view.");
                    return Result.Cancelled;
                }

                // Snapshot BEFORE
                HashSet<ElementId> before = CollectDimensionIdsInView(doc, view);

                // Input
                IList<FamilyInstance> sleeves = CollectSleevesInView(doc, view);
                if (sleeves.Count == 0)
                {
                    TaskDialog.Show(DialogTitle, "No sleeves (Conduit Fittings) found in this view.");
                    return Result.Cancelled;
                }

                // Find column grids (host + links). If none detected, fallback to all grids in view.
                IList<Grid> columnGrids = CollectColumnGridsHostAndLinks(doc, view);
                if (columnGrids.Count == 0)
                {
                    // fallback to all grids in the view (still gives you the behavior you asked for)
                    columnGrids = new FilteredElementCollector(doc, view.Id)
                        .OfClass(typeof(Grid))
                        .Cast<Grid>()
                        .ToList();

                    if (columnGrids.Count == 0)
                    {
                        TaskDialog.Show(DialogTitle, "No grids found in this view.");
                        return Result.Cancelled;
                    }
                }

                var (verticalGrids, horizontalGrids) = SplitGridsByOrientation(columnGrids);

                int attempts = 0;
                int failed = 0;
                int skipped = 0;

                using (Transaction tx = new Transaction(doc, "Dimensions to Sleeves"))
                {
                    tx.Start();

                    foreach (FamilyInstance sleeve in sleeves)
                    {
                        XYZ p = TryGetSleevePoint(sleeve);
                        if (p == null) { skipped++; continue; }

                        Grid gVert = FindNearestGrid(p, verticalGrids);
                        Grid gHorz = FindNearestGrid(p, horizontalGrids);

                        if (gVert != null)
                        {
                            attempts++;
                            if (!TryCreateLinearDimension(doc, view, sleeve, p, gVert, horizontalDirection: true))
                                failed++;
                        }
                        else skipped++;

                        if (gHorz != null)
                        {
                            attempts++;
                            if (!TryCreateLinearDimension(doc, view, sleeve, p, gHorz, horizontalDirection: false))
                                failed++;
                        }
                        else skipped++;
                    }

                    tx.Commit();
                }

                // Snapshot AFTER
                HashSet<ElementId> after = CollectDimensionIdsInView(doc, view);
                var newIds = after.Where(id => !before.Contains(id)).ToList();

                int segments = 0;
                foreach (var id in newIds)
                {
                    Dimension d = doc.GetElement(id) as Dimension;
                    if (d == null) continue;
                    if (d.NumberOfSegments > 0 && d.Segments != null) segments += d.Segments.Size;
                    else segments += 1;
                }

                string report =
                    $"Placed {newIds.Count} new dimension element(s)\n" +
                    $"(total {segments} segment(s))." +
                    $"\n\nAttempts: {attempts}" +
                    $"\nSkipped: {skipped}" +
                    $"\nFailed: {failed}";

                TaskDialog.Show(DialogTitle, report);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show(DialogTitle, "Dimensions to Sleeves failed:\n\n" + ex.Message);
                return Result.Failed;
            }
        }

        // ---------------- collection ----------------

        private static IList<FamilyInstance> CollectSleevesInView(Document doc, View view)
        {
            return new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_ConduitFitting)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();
        }

        /// <summary>
        /// Collect grids from the view that touch columns in the host model or any visible linked model.
        /// Searches both Structural Columns and Architectural Columns.
        /// If nothing is found, caller will decide the fallback.
        /// </summary>
        private static IList<Grid> CollectColumnGridsHostAndLinks(Document doc, View view)
        {
            var grids = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .ToList();

            if (grids.Count == 0) return new List<Grid>();

            // Host model columns (structural + architectural)
            var hostColumns = new List<Element>();
            hostColumns.AddRange(new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_StructuralColumns)
                .WhereElementIsNotElementType()
                .ToElements());

            hostColumns.AddRange(new FilteredElementCollector(doc, view.Id)
                .OfCategory(BuiltInCategory.OST_Columns)
                .WhereElementIsNotElementType()
                .ToElements());

            // Columns from visible links (structural + architectural)
            var linkColumnsTransformed = new List<BoundingBoxXYZ>();
            var links = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            foreach (var link in links)
            {
                Document ldoc = link.GetLinkDocument();
                if (ldoc == null) continue;

                Transform t = link.GetTotalTransform();

                IEnumerable<Element> lcols1 = new FilteredElementCollector(ldoc)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .WhereElementIsNotElementType()
                    .ToElements();

                IEnumerable<Element> lcols2 = new FilteredElementCollector(ldoc)
                    .OfCategory(BuiltInCategory.OST_Columns)
                    .WhereElementIsNotElementType()
                    .ToElements();

                foreach (var col in lcols1.Concat(lcols2))
                {
                    BoundingBoxXYZ bb = col.get_BoundingBox(null);
                    if (bb == null) continue;

                    // bring into host coords
                    BoundingBoxXYZ tbb = new BoundingBoxXYZ
                    {
                        Min = t.OfPoint(bb.Min),
                        Max = t.OfPoint(bb.Max)
                    };
                    linkColumnsTransformed.Add(Grow(tbb, 0.5)); // expand ~6"
                }
            }

            bool HasAnyColumns()
                => hostColumns.Count > 0 || linkColumnsTransformed.Count > 0;

            if (!HasAnyColumns())
                return new List<Grid>(); // let caller fallback to all grids

            var result = new List<Grid>();

            foreach (var g in grids)
            {
                Curve gc = TryGetGridCurve(g);
                if (gc == null) continue;

                bool touchesAColumn = false;

                // Host columns
                foreach (var col in hostColumns)
                {
                    var bb = col.get_BoundingBox(view);
                    if (bb == null) continue;

                    var bbGrow = Grow(bb, 0.5); // ~6"
                    if (CurveIntersectsBox(gc, bbGrow))
                    {
                        touchesAColumn = true;
                        break;
                    }
                }

                // Link columns (already grown)
                if (!touchesAColumn)
                {
                    foreach (var tbb in linkColumnsTransformed)
                    {
                        if (CurveIntersectsBox(gc, tbb))
                        {
                            touchesAColumn = true;
                            break;
                        }
                    }
                }

                if (touchesAColumn) result.Add(g);
            }

            return result;
        }

        private static (IList<Grid> vertical, IList<Grid> horizontal) SplitGridsByOrientation(IList<Grid> grids)
        {
            var vertical = new List<Grid>();
            var horizontal = new List<Grid>();

            foreach (var g in grids)
            {
                Curve c = TryGetGridCurve(g);
                if (c == null) continue;

                XYZ d = (c.GetEndPoint(1) - c.GetEndPoint(0)).Normalize();
                if (Math.Abs(d.Y) >= Math.Abs(d.X))
                    vertical.Add(g);
                else
                    horizontal.Add(g);
            }
            return (vertical, horizontal);
        }

        private static Grid FindNearestGrid(XYZ p, IList<Grid> candidates)
        {
            double best = VeryLarge;
            Grid bestG = null;

            foreach (var g in candidates)
            {
                Curve c = TryGetGridCurve(g);
                if (c == null) continue;
                double d = c.Distance(p);
                if (d < best) { best = d; bestG = g; }
            }
            return bestG;
        }

        // ---------------- dimension creation ----------------

        private static bool TryCreateLinearDimension(
            Document doc, View view, FamilyInstance sleeve, XYZ sleevePt, Grid grid, bool horizontalDirection)
        {
            try
            {
                Reference rGrid;
                if (!TryGetGridReference(grid, out rGrid))
                    return false;

                Reference rSleeve;
                if (!TryGetSleeveReference(doc, sleeve, horizontalDirection, out rSleeve))
                    return false;

                XYZ p0, p1;
                if (horizontalDirection)
                {
                    p0 = new XYZ(sleevePt.X - 1000, sleevePt.Y, 0);
                    p1 = new XYZ(sleevePt.X + 1000, sleevePt.Y, 0);
                }
                else
                {
                    p0 = new XYZ(sleevePt.X, sleevePt.Y - 1000, 0);
                    p1 = new XYZ(sleevePt.X, sleevePt.Y + 1000, 0);
                }

                Line dimLine = Line.CreateBound(p0, p1);

                var refs = new ReferenceArray();
                refs.Append(rGrid);
                refs.Append(rSleeve);

                Dimension dim = doc.Create.NewDimension(view, dimLine, refs);
                return dim != null;
            }
            catch
            {
                return false;
            }
        }

        // Revit 2024-safe: use a datum reference to the grid
        private static bool TryGetGridReference(Grid g, out Reference r)
        {
            r = null;
            try
            {
                r = new Reference(g);
                return r != null;
            }
            catch
            {
                r = null;
                return false;
            }
        }

        // Prefer family center references; fallback to face aligned with required direction
        private static bool TryGetSleeveReference(Document doc, FamilyInstance fi, bool horizontalDirection, out Reference r)
        {
            r = null;

            try
            {
                FamilyInstanceReferenceType want = horizontalDirection
                    ? FamilyInstanceReferenceType.CenterLeftRight
                    : FamilyInstanceReferenceType.CenterFrontBack;

                IList<Reference> centers = fi.GetReferences(want);
                if (centers != null && centers.Count > 0)
                {
                    r = centers[0];
                    return true;
                }
            }
            catch { }

            try
            {
                Options opt = new Options { ComputeReferences = true, DetailLevel = ViewDetailLevel.Fine };
                GeometryElement ge = fi.get_Geometry(opt);
                if (ge != null)
                {
                    foreach (GeometryObject go in ge)
                    {
                        Solid s = go as Solid;
                        if (s == null || s.Faces == null) continue;

                        foreach (Face f in s.Faces)
                        {
                            if (f.Area < MinFaceArea) continue;

                            BoundingBoxUV bb = f.GetBoundingBox();
                            UV mid = new UV((bb.Min.U + bb.Max.U) * 0.5, (bb.Min.V + bb.Max.V) * 0.5);
                            XYZ n = f.ComputeNormal(mid);
                            if (n == null) continue;

                            n = new XYZ(Math.Abs(n.X), Math.Abs(n.Y), Math.Abs(n.Z));
                            bool ok =
                                (horizontalDirection && n.X > n.Y + 1e-4 && n.X > n.Z + 1e-4) ||
                                (!horizontalDirection && n.Y > n.X + 1e-4 && n.Y > n.Z + 1e-4);

                            if (ok && f.Reference != null)
                            {
                                r = f.Reference;
                                return true;
                            }
                        }
                    }
                }
            }
            catch { }

            return false;
        }

        // ---------------- small helpers ----------------

        private static XYZ TryGetSleevePoint(FamilyInstance fi)
        {
            try
            {
                var lp = fi.Location as LocationPoint;
                if (lp != null) return lp.Point;

                BoundingBoxXYZ bb = fi.get_BoundingBox(null);
                if (bb != null) return (bb.Min + bb.Max) * 0.5;

                return null;
            }
            catch { return null; }
        }

        private static Curve TryGetGridCurve(Grid g)
        {
            try
            {
                return g.Curve;
            }
            catch { return null; }
        }

        private static bool CurveIntersectsBox(Curve c, BoundingBoxXYZ box)
        {
            if (c == null || box == null) return false;

            var p0 = c.GetEndPoint(0);
            var p1 = c.GetEndPoint(1);

            if (PointInBox(p0, box) || PointInBox(p1, box)) return true;

            var line = Line.CreateBound(p0, p1);
            if (line == null) return false;

            XYZ[] samples = new[]
            {
                new XYZ(box.Min.X, box.Min.Y, 0),
                new XYZ(box.Min.X, box.Max.Y, 0),
                new XYZ(box.Max.X, box.Min.Y, 0),
                new XYZ(box.Max.X, box.Max.Y, 0),
                new XYZ((box.Min.X+box.Max.X)/2.0, (box.Min.Y+box.Max.Y)/2.0, 0)
            };

            foreach (var s in samples)
            {
                double d = DistancePointToLineXY(s, line);
                if (d <= 0.5) return true; // ~6"
            }
            return false;
        }

        private static bool PointInBox(XYZ p, BoundingBoxXYZ b)
        {
            return p.X >= b.Min.X - Tolerance && p.X <= b.Max.X + Tolerance
                && p.Y >= b.Min.Y - Tolerance && p.Y <= b.Max.Y + Tolerance
                && p.Z >= b.Min.Z - Tolerance && p.Z <= b.Max.Z + Tolerance;
        }

        private static double DistancePointToLineXY(XYZ p, Line line)
        {
            XYZ a = line.GetEndPoint(0);
            XYZ b = line.GetEndPoint(1);
            XYZ ap = new XYZ(p.X - a.X, p.Y - a.Y, 0);
            XYZ ab = new XYZ(b.X - a.X, b.Y - a.Y, 0);

            double ab2 = ab.X * ab.X + ab.Y * ab.Y;
            if (ab2 < Tolerance) return ap.GetLength();

            double t = (ap.X * ab.X + ap.Y * ab.Y) / ab2;
            XYZ proj = new XYZ(a.X + t * ab.X, a.Y + t * ab.Y, 0);
            return (new XYZ(p.X - proj.X, p.Y - proj.Y, 0)).GetLength();
        }

        private static BoundingBoxXYZ Grow(BoundingBoxXYZ bb, double ft)
        {
            return new BoundingBoxXYZ
            {
                Min = new XYZ(bb.Min.X - ft, bb.Min.Y - ft, bb.Min.Z - ft),
                Max = new XYZ(bb.Max.X + ft, bb.Max.Y + ft, bb.Max.Z + ft)
            };
        }

        private static HashSet<ElementId> CollectDimensionIdsInView(Document doc, View view)
        {
            var set = new HashSet<ElementId>();
            var it = new FilteredElementCollector(doc, view.Id)
                .OfClass(typeof(Dimension))
                .GetElementIterator();

            it.Reset();
            while (it.MoveNext())
            {
                Element e = it.Current as Element;
                if (e != null) set.Add(e.Id);
            }
            return set;
        }
    }
}
