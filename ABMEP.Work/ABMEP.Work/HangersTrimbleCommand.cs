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

namespace ABMEP.Work
{
    [Transaction(TransactionMode.Manual)]
    public class HangersTrimbleCommand : IExternalCommand
    {
        private const string BUILD_TAG = "[HT-rod v19]";

        // Settings
        private const bool USE_SHARED_COORDS = false;
        private const string OUTPUT_DIR = @"C:\Temp";
        private const string DEFAULT_LEVEL = "Level";

        // Geometry thresholds
        private const double VERTICAL_DOT_MIN = 0.70;       // |uz| ≥ 0.70 → vertical-ish
        private const double MIN_ROD_ZSPAN_FT = 0.04;       // ~0.5"
        private const int EDGE_TESSELATION_MAX = 100;
        private const double CYL_VERTICAL_DOT_MIN = 0.97;       // cylinder axis ~ vertical
        private const double CIRCLE_PLANE_DOT_MIN = 0.97;       // arc plane normal ~ Z

        // Unistrut fallback heuristics
        private const double UNISTRUT_RATIO_MIN = 1.6;        // long/short ≥ 1.6
        private const double UNISTRUT_LONG_MIN_FT = 0.50;       // long side ≥ 6"

        // Duplicate suppression
        private const double NEAR_DUPLICATE_TOL_FT = 1.5 / 12.0; // 1.5"
        private const double ROD_SMALL_DEDUP_TOL_FT = 0.20 / 12.0;// 0.20"

        // Rod-by-BBox (weak hint)
        private const double ROD_MAX_DIAM_FT = 1.5 / 12.0; // ≤ 1.5"
        private const double ROD_MIN_HEIGHT_FT = 3.0 / 12.0; // ≥ 3"
        private const double ROD_SLENDER_RATIO = 4.0;        // height / max(dx,dy) ≥ 4

        // Edge->rod (legacy bundler) – fallback
        private const double EDGE_CLUSTER_TOL_FT = 0.02;       // ~1/4"
        private const double ROD_BUNDLE_TOL_FT = 0.075;      // ~15/16"
        private const int ROD_MIN_EDGES_IN_BUNDLE = 3;

        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = data.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                if (doc == null || doc.ActiveView == null)
                {
                    message = "No active document/view.";
                    return Result.Cancelled;
                }

                Transform world = USE_SHARED_COORDS
                    ? (doc.ActiveProjectLocation != null ? doc.ActiveProjectLocation.GetTotalTransform() : Transform.Identity)
                    : Transform.Identity;

                var hangers = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_FabricationHangers)
                    .WhereElementIsNotElementType()
                    .OfType<FabricationPart>()
                    .ToList();

                if (hangers.Count == 0)
                {
                    TaskDialog.Show($"Hangers Trimble {BUILD_TAG}", "No MEP Fabrication Hangers found in the active view.");
                    return Result.Cancelled;
                }

                var singles = new List<HPoint>();
                var doublesA = new List<HPoint>();
                var doublesB = new List<HPoint>();
                var skipped = new List<SkipRow>();

                foreach (var fp in hangers)
                {
                    // Pull formatted diameter strings in order (rod_1, rod_2, ...)
                    var diaTxt = GetRodDiameterStrings(fp);

                    // Strong override for clevis/band: exactly 1 top-center
                    if (IsClevisOrBand(fp, doc))
                    {
                        var tc = TopCenterFromBB(fp, doc, world);
                        if (tc != null)
                        {
                            singles.Add(new HPoint
                            {
                                ElemId = fp.Id,
                                P = tc,
                                Description = AppendDia("Hanger", diaTxt, 0)
                            });
                            continue;
                        }
                    }

                    // 1) Use cylinder-by-diameter FIRST
                    var tops = RodTopsFromCylindersByDiameter(fp, doc, world);

                    // 2) If nothing, try connectors
                    if (tops.Count == 0)
                        tops = RodTopsFromConnectors(fp, world);

                    // 3) If still nothing, try geometry edge bundler + bbox helper
                    if (tops.Count == 0)
                        tops = RodTopsFromGeometryFallback(fp, doc, world);

                    // 4) Last resort: unistrut ends or center from BB
                    if (tops.Count == 0)
                        tops = BBFallbackTops(fp, doc, world);

                    // Final de-dup
                    tops = CollapseNearDuplicates(tops, NEAR_DUPLICATE_TOL_FT);

                    if (tops.Count == 1)
                    {
                        singles.Add(new HPoint
                        {
                            ElemId = fp.Id,
                            P = tops[0],
                            Description = AppendDia("Hanger", diaTxt, 0)
                        });
                    }
                    else if (tops.Count >= 2)
                    {
                        var pair = FarthestTwoXY(tops);
                        pair.Sort((p1, p2) => p1.X != p2.X ? p1.X.CompareTo(p2.X) : p1.Y.CompareTo(p2.Y));

                        doublesA.Add(new HPoint
                        {
                            ElemId = fp.Id,
                            P = pair[0],
                            Description = AppendDia("Hanger", diaTxt, 0)
                        });
                        doublesB.Add(new HPoint
                        {
                            ElemId = fp.Id,
                            P = pair[1],
                            Description = AppendDia("Hanger", diaTxt, 1)
                        });
                    }
                    else
                    {
                        skipped.Add(new SkipRow { ElementId = fp.Id.Value, Reason = "No rod points detected" });
                    }
                }

                if (singles.Count == 0 && doublesA.Count == 0)
                {
                    TaskDialog.Show($"Hangers Trimble {BUILD_TAG}", "No 1-rod or 2-rod hangers resolved in the active view.");
                    return Result.Cancelled;
                }

                // File names
                string proj = (doc.ProjectInformation != null ? doc.ProjectInformation.Name : "Project") ?? "Project";
                proj = proj.Trim();
                string lvl = PromptLevel(DEFAULT_LEVEL);
                if (string.IsNullOrWhiteSpace(lvl)) lvl = DEFAULT_LEVEL;
                string date = DateTime.Now.ToString("yyyy-MM-dd");
                string root = San(proj) + "_" + San(lvl) + "_Hanger Trimble_" + date;

                Directory.CreateDirectory(OUTPUT_DIR);

                // 1-rod
                if (singles.Count > 0)
                {
                    using (var sw = NewWriter(Path.Combine(OUTPUT_DIR, root + "_1Rod.csv")))
                    {
                        sw.WriteLine("Name,X,Y,Z,Description");
                        int n = 1;
                        foreach (var hp in singles) WriteRow(sw, "C-" + (n++), hp.P, hp.Description);
                    }
                    using (var sw = NewWriter(Path.Combine(OUTPUT_DIR, root + "_1Rod_ID.csv")))
                    {
                        sw.WriteLine("Name,X,Y,Z,Description,Element ID");
                        int n = 1;
                        foreach (var hp in singles) WriteRowWithId(sw, "C-" + (n++), hp.P, hp.Description, hp.ElemId.Value);
                    }
                }

                // 2-rod
                if (doublesA.Count > 0)
                {
                    using (var sw = NewWriter(Path.Combine(OUTPUT_DIR, root + "_2Rod.csv")))
                    {
                        sw.WriteLine("Name,X,Y,Z,Description");
                        int u = 1;
                        for (int i = 0; i < doublesA.Count; ++i)
                        {
                            WriteRow(sw, "U-" + (u++), doublesA[i].P, doublesA[i].Description);
                            WriteRow(sw, "U-" + (u++), doublesB[i].P, doublesB[i].Description);
                        }
                    }
                    using (var sw = NewWriter(Path.Combine(OUTPUT_DIR, root + "_2Rod_ID.csv")))
                    {
                        sw.WriteLine("Name,X,Y,Z,Description,Element ID");
                        int u = 1;
                        for (int i = 0; i < doublesA.Count; ++i)
                        {
                            WriteRowWithId(sw, "U-" + (u++), doublesA[i].P, doublesA[i].Description, doublesA[i].ElemId.Value);
                            WriteRowWithId(sw, "U-" + (u++), doublesB[i].P, doublesB[i].Description, doublesB[i].ElemId.Value);
                        }
                    }
                }

                if (skipped.Count > 0)
                {
                    using (var sw = NewWriter(Path.Combine(OUTPUT_DIR, root + "_Skipped.csv")))
                    {
                        sw.WriteLine("ElementId,Reason");
                        foreach (var s in skipped)
                            sw.WriteLine(s.ElementId.ToString(CultureInfo.InvariantCulture) + "," + Csv(s.Reason));
                    }
                }

                TaskDialog.Show(
                    "Hangers Trimble " + BUILD_TAG,
                    "Export complete.\n" +
                    (singles.Count > 0 ? "  1-rod: " + singles.Count + "\n" : "") +
                    (doublesA.Count > 0 ? "  2-rod: " + doublesA.Count + " hangers (" + (doublesA.Count * 2) + " points)\n" : "") +
                    (skipped.Count > 0 ? "  Skipped: " + skipped.Count + " (see *_Skipped.csv)\n" : "")
                );

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = BUILD_TAG + " " + ex;
                return Result.Failed;
            }
        }

        // ===== Primary: cylinder-by-diameter (unchanged logic) =====

        private static List<XYZ> RodTopsFromCylindersByDiameter(FabricationPart fp, Document doc, Transform world)
        {
            var tops = new List<XYZ>();

            var expectedDias = GetRodDiametersFeetOrdered(fp); // ordered if possible
            if (expectedDias.Count == 0)
                expectedDias = GetRodDiametersFeetFallback();   // common sizes

            try
            {
                var opt = new Options { ComputeReferences = false, IncludeNonVisibleObjects = true, DetailLevel = ViewDetailLevel.Fine };
                var ge = fp.get_Geometry(opt);
                if (ge == null) return tops;

                foreach (var go in ge)
                {
                    var gi = go as GeometryInstance;
                    if (gi != null)
                    {
                        var sub = gi.GetInstanceGeometry();
                        if (sub == null) continue;
                        CollectCylRodTops(sub, gi.Transform, world, expectedDias, tops);
                    }
                    else
                    {
                        CollectCylRodTops(go, Transform.Identity, world, expectedDias, tops);
                    }
                }
            }
            catch { }

            return CollapseNearDuplicates(tops, ROD_SMALL_DEDUP_TOL_FT);
        }

        private static void CollectCylRodTops(object geom, Transform acc, Transform world, List<double> expectedDias, List<XYZ> tops)
        {
            var ge = geom as GeometryElement;
            if (ge == null) return;

            foreach (var obj in ge)
            {
                var gi = obj as GeometryInstance;
                if (gi != null)
                {
                    var sub = gi.GetInstanceGeometry();
                    if (sub != null) CollectCylRodTops(sub, acc.Multiply(gi.Transform), world, expectedDias, tops);
                    continue;
                }

                var s = obj as Solid;
                if (s == null || s.Faces.Size == 0) continue;

                foreach (Face f in s.Faces)
                {
                    var cyl = f.GetSurface() as CylindricalSurface;
                    if (cyl == null) continue;

                    var axis = cyl.Axis; if (axis == null) continue;
                    axis = axis.Normalize();
                    if (Math.Abs(axis.DotProduct(XYZ.BasisZ)) < CYL_VERTICAL_DOT_MIN) continue;

                    double dia = 2.0 * cyl.Radius; // feet
                    if (!MatchesAnyDiameter(dia, expectedDias)) continue;

                    var top = GetHighestCircularEdgeCenter(f);
                    if (top != null)
                        tops.Add(world.OfPoint(acc.OfPoint(top)));
                }
            }
        }

        private static bool MatchesAnyDiameter(double diaFt, List<double> expectedDias)
        {
            const double tol = 0.03 / 12.0; // ~0.03" tolerance in feet
            foreach (var d in expectedDias)
                if (Math.Abs(diaFt - d) <= tol) return true;
            return false;
        }

        // ===== Secondary: connectors =====

        private static List<XYZ> RodTopsFromConnectors(FabricationPart fp, Transform world)
        {
            var list = new List<XYZ>();
            try
            {
                var cons = fp != null && fp.ConnectorManager != null ? fp.ConnectorManager.Connectors : null;
                if (cons != null)
                {
                    foreach (Connector c in cons)
                    {
                        if (c == null) continue;
                        XYZ axis = null;
                        try { axis = c.CoordinateSystem != null ? c.CoordinateSystem.BasisZ : null; } catch { }
                        axis = SafeNorm(axis);
                        if (axis == null) continue;
                        if (Math.Abs(axis.Z) >= VERTICAL_DOT_MIN)
                            list.Add(world.OfPoint(c.Origin));
                    }
                }
            }
            catch { }
            return CollapseNearDuplicates(list, ROD_SMALL_DEDUP_TOL_FT);
        }

        // ===== Tertiary: geometry-based fallback (unchanged) =====

        private static List<XYZ> RodTopsFromGeometryFallback(FabricationPart fp, Document doc, Transform world)
        {
            var edgePts = new List<XYZ>();
            try
            {
                var opt = new Options { ComputeReferences = false, IncludeNonVisibleObjects = true, DetailLevel = ViewDetailLevel.Fine };
                var ge = fp.get_Geometry(opt);
                if (ge != null)
                {
                    HarvestVerticalEdgeSamples(ge, Transform.Identity, world, edgePts);
                    HarvestRodTopsByBBox(ge, Transform.Identity, world, edgePts);
                    HarvestCylindricalRodTops(ge, Transform.Identity, world, edgePts);
                }
            }
            catch { }

            return BuildRodsFromEdgeClusters(edgePts);
        }

        private class EdgeCluster { public double Cx, Cy, ZMin, ZMax; }

        private static List<XYZ> BuildRodsFromEdgeClusters(List<XYZ> samples)
        {
            var result = new List<XYZ>();
            if (samples == null || samples.Count == 0) return result;

            var edgeGroups = ClusterXY(samples, EDGE_CLUSTER_TOL_FT);

            var edges = new List<EdgeCluster>(edgeGroups.Count);
            foreach (var g in edgeGroups)
            {
                double zmin = double.MaxValue, zmax = double.MinValue, sx = 0, sy = 0; int n = 0;
                foreach (var p in g)
                {
                    if (p.Z < zmin) zmin = p.Z; if (p.Z > zmax) zmax = p.Z;
                    sx += p.X; sy += p.Y; n++;
                }
                if ((zmax - zmin) < MIN_ROD_ZSPAN_FT) continue;
                edges.Add(new EdgeCluster { Cx = sx / Math.Max(1, n), Cy = sy / Math.Max(1, n), ZMin = zmin, ZMax = zmax });
            }
            if (edges.Count == 0) return result;

            var used = new bool[edges.Count];
            for (int i = 0; i < edges.Count; i++)
            {
                if (used[i]) continue;
                var bundle = new List<int> { i };
                used[i] = true;

                for (int j = i + 1; j < edges.Count; j++)
                {
                    if (used[j]) continue;
                    double d = Math.Sqrt(Math.Pow(edges[i].Cx - edges[j].Cx, 2) + Math.Pow(edges[i].Cy - edges[j].Cy, 2));
                    if (d <= ROD_BUNDLE_TOL_FT) { used[j] = true; bundle.Add(j); }
                }

                if (bundle.Count >= ROD_MIN_EDGES_IN_BUNDLE)
                {
                    double sx = 0, sy = 0, zTop = double.MinValue;
                    foreach (var idx in bundle)
                    {
                        sx += edges[idx].Cx; sy += edges[idx].Cy;
                        if (edges[idx].ZMax > zTop) zTop = edges[idx].ZMax;
                    }
                    result.Add(new XYZ(sx / bundle.Count, sy / bundle.Count, zTop));
                }
            }
            return result;
        }

        private static void HarvestVerticalEdgeSamples(GeometryElement ge, Transform acc, Transform world, List<XYZ> outPts)
        {
            foreach (var go in ge)
            {
                var gi = go as GeometryInstance;
                if (gi != null)
                {
                    var t = acc.Multiply(gi.Transform);
                    var sub = gi.GetInstanceGeometry();
                    if (sub != null) HarvestVerticalEdgeSamples(sub, t, world, outPts);
                    continue;
                }

                var s = go as Solid;
                if (s != null && s.Volume > 1e-9)
                {
                    foreach (Edge e in s.Edges)
                    {
                        Curve cv; try { cv = e.AsCurve(); } catch { continue; }
                        if (cv == null) continue;

                        var ln = cv as Line;
                        if (ln != null)
                        {
                            var d = SafeNorm(ln.Direction);
                            if (d != null && Math.Abs(d.Z) >= 0.95)
                            {
                                outPts.Add(world.OfPoint(acc.OfPoint(ln.GetEndPoint(0))));
                                outPts.Add(world.OfPoint(acc.OfPoint(ln.GetEndPoint(1))));
                                continue;
                            }
                        }

                        try
                        {
                            var tss = cv.Tessellate();
                            if (tss != null && tss.Count >= 2)
                            {
                                int step = Math.Max(1, tss.Count / EDGE_TESSELATION_MAX);
                                for (int i = 0; i + step < tss.Count; i += step)
                                {
                                    var a = acc.OfPoint(tss[i]);
                                    var b = acc.OfPoint(tss[i + step]);
                                    var d = SafeNorm(b - a);
                                    if (d != null && Math.Abs(d.Z) >= 0.95)
                                    {
                                        outPts.Add(world.OfPoint(a));
                                        outPts.Add(world.OfPoint(b));
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        // Weak bbox-based hints (feeds bundler)
        private static void HarvestRodTopsByBBox(GeometryElement ge, Transform acc, Transform world, List<XYZ> outPts)
        {
            foreach (var go in ge)
            {
                var gi = go as GeometryInstance;
                if (gi != null)
                {
                    var t = acc.Multiply(gi.Transform);
                    var sub = gi.GetInstanceGeometry();
                    if (sub != null) HarvestRodTopsByBBox(sub, t, world, outPts);
                    continue;
                }

                var s = go as Solid;
                if (s == null || s.Edges == null || s.Edges.Size == 0) continue;

                double minX = double.PositiveInfinity, minY = double.PositiveInfinity, minZ = double.PositiveInfinity;
                double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity, maxZ = double.NegativeInfinity;

                foreach (Edge e in s.Edges)
                {
                    try
                    {
                        var cv = e.AsCurve();
                        var ts = cv.Tessellate();
                        if (ts == null || ts.Count == 0) continue;
                        foreach (var p in ts)
                        {
                            var q = acc.OfPoint(p);
                            if (q.X < minX) minX = q.X; if (q.X > maxX) maxX = q.X;
                            if (q.Y < minY) minY = q.Y; if (q.Y > maxY) maxY = q.Y;
                            if (q.Z < minZ) minZ = q.Z; if (q.Z > maxZ) maxZ = q.Z;
                        }
                    }
                    catch { }
                }

                if (!IsFiniteCompat(minX) || !IsFiniteCompat(maxX) ||
                    !IsFiniteCompat(minY) || !IsFiniteCompat(maxY) ||
                    !IsFiniteCompat(minZ) || !IsFiniteCompat(maxZ))
                    continue;

                double dx = maxX - minX;
                double dy = maxY - minY;
                double dz = maxZ - minZ;
                double maxXY = Math.Max(dx, dy);
                if (dz >= ROD_MIN_HEIGHT_FT && maxXY <= ROD_MAX_DIAM_FT && (dz / Math.Max(maxXY, 1e-6)) >= ROD_SLENDER_RATIO)
                {
                    double cx = 0.5 * (minX + maxX);
                    double cy = 0.5 * (minY + maxY);
                    outPts.Add(world.OfPoint(acc.OfPoint(new XYZ(cx, cy, maxZ))));
                    outPts.Add(world.OfPoint(acc.OfPoint(new XYZ(cx, cy, 0.5 * (minZ + maxZ)))));
                }
            }
        }

        private static void HarvestCylindricalRodTops(GeometryElement ge, Transform acc, Transform world, List<XYZ> outPts)
        {
            foreach (var go in ge)
            {
                var gi = go as GeometryInstance;
                if (gi != null)
                {
                    var t = acc.Multiply(gi.Transform);
                    var sub = gi.GetInstanceGeometry();
                    if (sub != null) HarvestCylindricalRodTops(sub, t, world, outPts);
                    continue;
                }

                var s = go as Solid;
                if (s == null || s.Faces.Size == 0) continue;

                foreach (Face f in s.Faces)
                {
                    var cyl = f.GetSurface() as CylindricalSurface;
                    if (cyl == null) continue;

                    var axis = cyl.Axis; if (axis == null) continue;
                    axis = axis.Normalize();
                    if (Math.Abs(axis.DotProduct(XYZ.BasisZ)) < CYL_VERTICAL_DOT_MIN) continue;

                    var top = GetHighestCircularEdgeCenter(f);
                    if (top != null)
                        outPts.Add(world.OfPoint(acc.OfPoint(top)));
                }
            }
        }

        private static XYZ GetHighestCircularEdgeCenter(Face face)
        {
            double topZ = double.NegativeInfinity;
            XYZ best = null;

            var loops = face.EdgeLoops;
            foreach (EdgeArray loop in loops)
            {
                foreach (Edge edge in loop)
                {
                    var c = edge.AsCurve();
                    var arc = c as Arc;
                    if (arc == null || !arc.IsBound) continue;

                    var n = arc.Normal; if (n == null) continue;
                    n = n.Normalize();
                    if (Math.Abs(n.DotProduct(XYZ.BasisZ)) < CIRCLE_PLANE_DOT_MIN) continue;

                    double z = arc.Center.Z;
                    if (z > topZ) { topZ = z; best = arc.Center; }
                }
            }
            return best;
        }

        // ===== Fallbacks =====

        private static List<XYZ> BBFallbackTops(FabricationPart fp, Document doc, Transform world)
        {
            var bb = fp.get_BoundingBox(null) ?? fp.get_BoundingBox(doc.ActiveView);
            if (bb == null) return new List<XYZ>();

            var min = world.OfPoint(bb.Min);
            var max = world.OfPoint(bb.Max);
            double zTop = Math.Max(min.Z, max.Z);

            double dx = Math.Abs(max.X - min.X);
            double dy = Math.Abs(max.Y - min.Y);
            double longSide = Math.Max(dx, dy);
            double shortSide = Math.Max(1e-6, Math.Min(dx, dy));
            double aspect = longSide / shortSide;

            double cx = 0.5 * (min.X + max.X);
            double cy = 0.5 * (min.Y + max.Y);

            bool looksUnistrut = LooksLikeUnistrut(fp, doc) || (aspect >= UNISTRUT_RATIO_MIN && longSide >= UNISTRUT_LONG_MIN_FT);
            if (looksUnistrut)
            {
                if (dx >= dy) return new List<XYZ> { new XYZ(min.X, cy, zTop), new XYZ(max.X, cy, zTop) };
                else return new List<XYZ> { new XYZ(cx, min.Y, zTop), new XYZ(cx, max.Y, zTop) };
            }
            return new List<XYZ> { new XYZ(cx, cy, zTop) };
        }

        private static bool LooksLikeUnistrut(FabricationPart fp, Document doc)
        {
            string text = BuildSearchCorpus(fp, doc);
            if (string.IsNullOrWhiteSpace(text)) return false;

            string[] keys = { "unistrut", "uni-strut", "trapeze", "channel", "strut", "kindorf", "kd" };
            foreach (string k in keys)
                if (text.Contains(k)) return true;
            return false;
        }

        // ===== CLEVIS/BAND helpers =====

        private static bool IsClevisOrBand(FabricationPart fp, Document doc)
        {
            string text = BuildSearchCorpus(fp, doc);
            if (string.IsNullOrWhiteSpace(text)) return false;

            string[] keys = { "clevis", "band", "strap", "u-bolt", "u bolt",
                              "pipe hanger", "rod hanger", "band hanger", "strap hanger" };
            foreach (string k in keys)
                if (text.Contains(k)) return true;

            return false;
        }

        private static string BuildSearchCorpus(FabricationPart fp, Document doc)
        {
            var sb = new StringBuilder();

            try { if (!string.IsNullOrWhiteSpace(fp.Name)) sb.Append(" ").Append(fp.Name); } catch { }

            try
            {
                ElementId tid = fp.GetTypeId();
                if (tid != null && tid != ElementId.InvalidElementId)
                {
                    Element t = doc.GetElement(tid);
                    if (t != null && !string.IsNullOrWhiteSpace(t.Name))
                        sb.Append(" ").Append(t.Name);
                }
            }
            catch { }

            try
            {
                foreach (Parameter p in fp.Parameters)
                {
                    if (p != null && p.StorageType == StorageType.String && p.HasValue)
                    {
                        string v = p.AsString();
                        if (!string.IsNullOrWhiteSpace(v)) sb.Append(" ").Append(v);
                    }
                }
            }
            catch { }

            return sb.ToString().ToLowerInvariant();
        }

        private static XYZ TopCenterFromBB(FabricationPart fp, Document doc, Transform world)
        {
            var bb = fp.get_BoundingBox(null) ?? fp.get_BoundingBox(doc.ActiveView);
            if (bb == null) return null;
            var min = world.OfPoint(bb.Min);
            var max = world.OfPoint(bb.Max);
            double zTop = Math.Max(min.Z, max.Z);
            double cx = 0.5 * (min.X + max.X);
            double cy = 0.5 * (min.Y + max.Y);
            return new XYZ(cx, cy, zTop);
        }

        // ===== Diameter helpers (NEW) =====

        private static List<double> GetRodDiametersFeetOrdered(FabricationPart fp)
        {
            // Try to read rod_1_diameter, rod_2_diameter, ... in order
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
                        valFeet = p.AsDouble();
                    }
                    else if (p.StorageType == StorageType.String)
                    {
                        string s = p.AsString();
                        double vf; if (TryParseFeetInches(s, out vf)) valFeet = vf;
                    }

                    if (!IsFiniteCompat(valFeet) || valFeet <= 0) continue;

                    // find index like ...rod_1_diameter...
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
            // look for "rod_1", "rod 2", "rod-3"
            for (int i = 1; i <= 8; i++)
            {
                if (nameLower.Contains("rod_" + i) || nameLower.Contains("rod " + i) || nameLower.Contains("rod-" + i))
                    return i;
            }
            return -1;
        }

        private static List<double> GetRodDiametersFeetFallback()
        {
            double In(double inches) => inches / 12.0;
            return new List<double> { In(0.375), In(0.5), In(0.625), In(0.75) };
        }

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

        private static string AppendDia(string baseText, List<string> dias, int index)
        {
            if (dias != null && index >= 0 && index < dias.Count && !string.IsNullOrWhiteSpace(dias[index]))
                return baseText + " - " + dias[index];
            else if (dias != null && dias.Count > 0 && !string.IsNullOrWhiteSpace(dias[0]) && index > 0)
                return baseText + " - " + dias[0]; // fallback to first if missing a second
            return baseText;
        }

        private static string InchesToNiceFraction(double inches)
        {
            // Round to nearest 1/16"
            double frac = Math.Round(inches * 16.0) / 16.0;
            int whole = (int)Math.Floor(frac);
            double rem = frac - whole;

            int num = (int)Math.Round(rem * 16.0);
            int den = 16;

            if (num == 0) return whole.ToString(CultureInfo.InvariantCulture) + "\"";

            // reduce fraction
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

        // Parse strings like 0' 0 3/8" or 0' 3/4"
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
                            if (ab.Length == 2)
                            {
                                double a, b;
                                if (double.TryParse(ab[0], out a) && double.TryParse(ab[1], out b) && b != 0)
                                    inch += a / b;
                            }
                        }
                        else
                        {
                            double val;
                            if (double.TryParse(part, out val)) inch += val;
                        }
                    }
                }
                else
                {
                    double.TryParse(s, out inch);
                }
                feet = ft + inch / 12.0;
                return feet > 0;
            }
            catch { return false; }
        }

        // ===== Utilities (unchanged) =====

        private static bool IsFiniteCompat(double v) => !(double.IsNaN(v) || double.IsInfinity(v));

        private static List<XYZ> CollapseNearDuplicates(List<XYZ> pts, double tolFt)
        {
            if (pts == null || pts.Count <= 1) return pts ?? new List<XYZ>();
            var keep = new List<XYZ>(pts);
            for (int i = 0; i < keep.Count; ++i)
            {
                for (int j = i + 1; j < keep.Count; ++j)
                {
                    if (DistXY(keep[i], keep[j]) <= tolFt)
                    {
                        XYZ hi = keep[i].Z >= keep[j].Z ? keep[i] : keep[j];
                        keep[i] = hi;
                        keep.RemoveAt(j);
                        j--;
                    }
                }
            }
            return keep;
        }

        private static List<XYZ> FarthestTwoXY(List<XYZ> pts)
        {
            if (pts.Count <= 2) return new List<XYZ>(pts);
            double best = -1; XYZ A = pts[0], B = pts[1];
            for (int i = 0; i < pts.Count; ++i)
                for (int j = i + 1; j < pts.Count; ++j)
                {
                    double d = DistXY(pts[i], pts[j]);
                    if (d > best) { best = d; A = pts[i]; B = pts[j]; }
                }
            return new List<XYZ> { A, B };
        }

        private static List<XYZ> LimitToTwo(List<XYZ> pts)
        {
            if (pts == null) return new List<XYZ>();
            if (pts.Count <= 2) return new List<XYZ>(pts);
            return FarthestTwoXY(pts);
        }

        private static List<List<XYZ>> ClusterXY(List<XYZ> pts, double tol)
        {
            var groups = new List<List<XYZ>>();
            foreach (XYZ p in pts)
            {
                bool placed = false;
                foreach (List<XYZ> g in groups)
                {
                    if (DistXY(p, g[0]) <= tol) { g.Add(p); placed = true; break; }
                }
                if (!placed) groups.Add(new List<XYZ> { p });
            }
            return groups;
        }

        private static double DistXY(XYZ a, XYZ b)
        {
            double dx = a.X - b.X, dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static XYZ SafeNorm(XYZ v)
        {
            if (v == null) return null;
            double L = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (L < 1e-9) return null;
            return new XYZ(v.X / L, v.Y / L, v.Z / L);
        }

        private static StreamWriter NewWriter(string path)
        {
            var sw = new StreamWriter(path, false, Encoding.UTF8) { NewLine = "\r\n" };
            return sw;
        }

        private static void WriteRow(StreamWriter sw, string name, XYZ p, string desc)
        {
            sw.WriteLine(Csv(name) + "," + ToInv(p.X) + "," + ToInv(p.Y) + "," + ToInv(p.Z) + "," + Csv(desc));
        }

        private static void WriteRowWithId(StreamWriter sw, string name, XYZ p, string desc, long elemId)
        {
            sw.WriteLine(Csv(name) + "," + ToInv(p.X) + "," + ToInv(p.Y) + "," + ToInv(p.Z) + "," + Csv(desc) + "," + elemId.ToString(CultureInfo.InvariantCulture));
        }

        private static string ToInv(double d) => d.ToString("0.########", CultureInfo.InvariantCulture);

        private static string Csv(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        private static string San(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Untitled";
            var bad = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder(name.Length);
            foreach (var ch in name) sb.Append(Array.IndexOf(bad, ch) >= 0 ? '_' : ch);
            return sb.ToString().Trim();
        }

        private static string PromptLevel(string def)
        {
            try { return SimpleTextPrompt.Show("Hanger Trimble " + BUILD_TAG, "Enter Level text for file names:", def); }
            catch { return def; }
        }

        // Tiny WinForms prompt
        private class SimpleTextPrompt : System.Windows.Forms.Form
        {
            private readonly System.Windows.Forms.TextBox _tb;
            private readonly System.Windows.Forms.Button _ok, _cancel;

            private SimpleTextPrompt(string title, string message, string initial)
            {
                Text = title;
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                MinimizeBox = false; MaximizeBox = false;
                Width = 420; Height = 160;

                var lbl = new System.Windows.Forms.Label { Left = 12, Top = 12, Width = 380, Text = message };
                _tb = new System.Windows.Forms.TextBox { Left = 12, Top = 38, Width = 380, Text = initial ?? "" };
                _ok = new System.Windows.Forms.Button { Text = "OK", Left = 226, Width = 80, Top = 70, DialogResult = System.Windows.Forms.DialogResult.OK };
                _cancel = new System.Windows.Forms.Button { Text = "Cancel", Left = 312, Width = 80, Top = 70, DialogResult = System.Windows.Forms.DialogResult.Cancel };

                Controls.Add(lbl); Controls.Add(_tb); Controls.Add(_ok); Controls.Add(_cancel);
                AcceptButton = _ok; CancelButton = _cancel;
            }

            public static string Show(string title, string message, string initial)
            {
                using (var f = new SimpleTextPrompt(title, message, initial))
                {
                    var r = f.ShowDialog();
                    return r == System.Windows.Forms.DialogResult.OK ? f._tb.Text : initial;
                }
            }
        }

        private class HPoint
        {
            public ElementId ElemId;
            public XYZ P;
            public string Description;
        }

        private class SkipRow
        {
            public long ElementId;
            public string Reason;
        }
    }
}
