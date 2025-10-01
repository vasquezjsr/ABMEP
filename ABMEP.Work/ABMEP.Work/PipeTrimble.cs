// ABMEP.Work/PipeTrimble.cs
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Fabrication;
using Autodesk.Revit.UI;

namespace ABMEP.Work
{
    /// <summary>
    /// Fabrication Pipework T.I. points (plan-only) → Trimble CSV.
    /// - Active view only
    /// - Straights only for line geometry (fittings auto-excluded)
    /// - Intersections of extended straights (XY) using:
    ///     • 3D centerline proximity (≤ JOIN_CENTERLINE_TOL_FT) and
    ///     • real model connectivity (≤ 3 hops through fittings/shorts)
    ///   Includes SHORT straights to support corners.
    /// - Riser→Horizontal hits via true connectivity (≤ 2 hops)
    /// - Open ends ONLY from regular straights if no intersection within 18"
    /// - Z = BOP using OUTSIDE radius with slope correction
    /// - Dedupe nearby points (XY ≤ 3", Z ≤ 1/4")
    /// - CSV header: Name,X,Y,Z,Description (Description="BOP")
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class PipeTrimble : IExternalCommand
    {
        // ---------- SETTINGS ----------
        private const bool USE_SHARED_COORDS = false;            // shared vs project
        private static readonly string OUTPUT_DIR = @"C:\Temp";
        private const string FILENAME = "Trimble_Points.csv";

        // Geometry thresholds/tuning
        private const double EXT_FACTOR = 2.0;        // extend each end by (2×Diameter + EXTRA_IN)
        private const double EXTRA_IN = 3.0;          // extra inches on the extension

        private const double REG_MIN_PLAN_LEN_FT = 0.50;         // ≥ 6" used for intersections + open ends
        private const double SHORT_MIN_PLAN_LEN_FT = 0.001;        // ≥ ~0.012" used to help intersections/risers
        private const double VERT_PLAN_TOL_FT = 0.05;         // ≤ ~0.6" plan-length => vertical post
        private const double VERT_MIN_Z_FT = 0.10;         // ≥ ~1.2" vertical height
        private const double LINEPOINT_TOL_FT = 0.50;         // pipe–vertical XY tolerance
        private const double END_NEAR_TOL_FT = 1.5;          // 18" = 1.5 ft

        // Intersection acceptance (tight to avoid false crossings)
        private const double JOIN_CENTERLINE_TOL_FT = 0.0833333333; // 1" centerline-to-centerline in 3D

        // De-dup tolerances
        private const double MERGE_XY_TOL_FT = 0.25;               // 3"
        private const double MERGE_Z_TOL_FT = 0.0208333333;       // 1/4" in feet

        private const string CSV_HEADER = "Name,X,Y,Z,Description";
        private const string DESCRIPTION_TEXT = "BOP";
        // --------------------------------

        private const double TOL = 1e-9, EPS = 1e-6;

        public Result Execute(ExternalCommandData cdata, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = cdata.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                Transform xf = USE_SHARED_COORDS
                    ? (doc.ActiveProjectLocation?.GetTotalTransform() ?? Transform.Identity)
                    : Transform.Identity;

                // ----- collect Fabrication parts from active view -----
                var parts = new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .OfCategory(BuiltInCategory.OST_FabricationPipework)
                    .WhereElementIsNotElementType()
                    .OfType<FabricationPart>()
                    .ToList();

                // Build connectivity graph across ALL fabrication parts (straights + fittings)
                var adjacency = BuildAdjacency(parts);

                var pipesReg = new List<PipeRec>();  // regular straights (>= 6")
                var pipesShort = new List<PipeRec>();  // short straights (>= 0.012", < 6")
                var verts = new List<VertRec>();  // vertical posts
                var straightById = new Dictionary<ElementId, PipeRec>();

                foreach (var fp in parts)
                {
                    if (!IsStraight(fp)) continue;

                    var cs = GetConnectors(fp);
                    if (cs.Count < 2) continue;

                    // 3D connector points (shared or project)
                    var cpts = cs.Select(c => xf.OfPoint(c.Origin)).ToList();
                    if (!FarthestTwo(cpts, out XYZ c0, out XYZ c1)) continue;

                    var p0 = Flatten(c0);
                    var p1 = Flatten(c1);
                    double Lp = Dist2D(p0, p1);
                    double dz = Math.Abs(c1.Z - c0.Z);

                    // vertical “post”
                    if (Lp <= VERT_PLAN_TOL_FT && dz >= VERT_MIN_Z_FT)
                    {
                        verts.Add(new VertRec { X = p0.X, Y = p0.Y, C0 = c0, C1 = c1, ElemId = fp.Id });
                        continue; // don't treat posts as horizontals
                    }

                    // ---- Use OUTSIDE radius (feet) robustly ----
                    double radOutFt = GetOutsideRadiusFeet(fp, cs);

                    double diamIn = 2.0 * radOutFt * 12.0;
                    double extFt = (EXT_FACTOR * diamIn + EXTRA_IN) / 12.0;

                    Dir2D(p0, p1, out double ux, out double uy, out _);
                    var e0 = new XYZ(p0.X - extFt * ux, p0.Y - extFt * uy, 0.0);
                    var e1 = new XYZ(p1.X + extFt * ux, p1.Y + extFt * uy, 0.0);

                    // slope-correct vertical radius component (rZ = r_out * sqrt(1 - uz^2))
                    double rZ = ComputeVerticalRadiusComponent(c0, c1, radOutFt);

                    var rec = new PipeRec
                    {
                        P0 = p0,
                        P1 = p1,
                        E0 = e0,
                        E1 = e1,
                        C0 = c0,
                        C1 = c1,
                        Lp = Lp,
                        Ext = extFt,
                        RadOut = radOutFt,
                        RadVertical = rZ,
                        ElemId = fp.Id
                    };

                    if (Lp >= REG_MIN_PLAN_LEN_FT)
                        pipesReg.Add(rec);
                    else if (Lp >= SHORT_MIN_PLAN_LEN_FT)
                        pipesShort.Add(rec);

                    // index any straight (reg or short)
                    straightById[fp.Id] = rec;
                }

                if (straightById.Count == 0 && verts.Count == 0)
                {
                    message = "No usable Fabrication straights in the active view.";
                    return Result.Cancelled;
                }

                // Build list for intersection/riser tests (regular + short)
                var pipesAll = new List<PipeRec>(pipesReg.Count + pipesShort.Count);
                pipesAll.AddRange(pipesReg);
                pipesAll.AddRange(pipesShort);

                // ----- intersections & ends -----
                var points = new List<XYZ>();

                // pipe–pipe intersections:
                // require XY intersection, 3D centerline proximity, and real connectivity (≤ 3 hops)
                for (int i = 0; i < pipesAll.Count; ++i)
                {
                    var A = pipesAll[i];
                    for (int j = i + 1; j < pipesAll.Count; ++j)
                    {
                        var B = pipesAll[j];

                        // Quick XY test on extended lines
                        if (!SegSegIntersect2D(A.E0, A.E1, B.E0, B.E1,
                                               out double ix, out double iy, out double ta, out double tb))
                            continue;

                        double tA = MapT_ExtToT3D(ta, A.Lp, A.Ext);
                        double tB = MapT_ExtToT3D(tb, B.Lp, B.Ext);

                        // 3D centerline points at those stations
                        XYZ pA3 = new XYZ(
                            A.C0.X + tA * (A.C1.X - A.C0.X),
                            A.C0.Y + tA * (A.C1.Y - A.C0.Y),
                            A.C0.Z + tA * (A.C1.Z - A.C0.Z));
                        XYZ pB3 = new XYZ(
                            B.C0.X + tB * (B.C1.X - B.C0.X),
                            B.C0.Y + tB * (B.C1.Y - B.C0.Y),
                            B.C0.Z + tB * (B.C1.Z - B.C0.Z));

                        // Must be near in 3D to be a true joint
                        if (pA3.DistanceTo(pB3) > JOIN_CENTERLINE_TOL_FT)
                            continue;

                        // Also must be connected in the model (≤ 3 hops via fittings/shorts)
                        if (!AreConnected(adjacency, A.ElemId, B.ElemId, maxHops: 3))
                            continue;

                        double zA = BopZAtT(A, tA);
                        double zB = BopZAtT(B, tB);
                        points.Add(new XYZ(ix, iy, Math.Min(zA, zB)));
                    }
                }

                // Riser→Horizontal hits using connectivity (≤ 2 hops)
                foreach (var V in verts)
                {
                    var startId = V.ElemId;
                    var candidates = FindConnectedStraights(startId, adjacency, straightById, maxHops: 2);
                    if (candidates.Count == 0) continue;

                    // choose candidate with minimal perpendicular distance in XY
                    PipeRec best = null;
                    double bestD = double.MaxValue, bestT = 0.0;

                    foreach (var A in candidates)
                    {
                        double d = DistPtToSeg2D(V.X, V.Y, A.E0, A.E1, out double tExt);
                        if (d < bestD)
                        {
                            bestD = d; best = A; bestT = tExt;
                        }
                    }

                    if (best != null && bestD <= LINEPOINT_TOL_FT)
                    {
                        double tA = MapT_ExtToT3D(bestT, best.Lp, best.Ext);
                        points.Add(new XYZ(V.X, V.Y, BopZAtT(best, tA)));
                    }
                }

                // open ends: only from regular straights, only if no intersection within 18"
                foreach (var A in pipesReg)
                {
                    if (!NearAnyXY(points, A.P0, END_NEAR_TOL_FT))
                        points.Add(new XYZ(A.P0.X, A.P0.Y, BopZAtT(A, 0.0)));
                    if (!NearAnyXY(points, A.P1, END_NEAR_TOL_FT))
                        points.Add(new XYZ(A.P1.X, A.P1.Y, BopZAtT(A, 1.0)));
                }

                if (points.Count == 0)
                {
                    message = "No intersections or eligible open ends found.";
                    return Result.Cancelled;
                }

                // ----- de-duplicate nearby points (reduce noise) -----
                var merged = MergeNearby(points, MERGE_XY_TOL_FT, MERGE_Z_TOL_FT);

                // ----- CSV -----
                Directory.CreateDirectory(OUTPUT_DIR);
                string path = Path.Combine(OUTPUT_DIR, FILENAME);
                using (var sw = new StreamWriter(path, false, System.Text.Encoding.UTF8))
                {
                    sw.NewLine = "\r\n";
                    sw.WriteLine(CSV_HEADER);
                    for (int i = 0; i < merged.Count; ++i)
                    {
                        var p = merged[i];
                        sw.WriteLine($"P-{i + 1},{p.X:R},{p.Y:R},{p.Z:R},{Csv(DESCRIPTION_TEXT)}");
                    }
                }

                TaskDialog.Show("ABMEP - Trimble CSV",
                    $"Wrote {merged.Count} stake points (from {points.Count} raw):\n{path}");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
        }

        // ===== records =====
        private class PipeRec
        {
            public XYZ P0, P1;  // plan endpoints
            public XYZ E0, E1;  // extended plan endpoints
            public XYZ C0, C1;  // true 3D endpoints (for Z)
            public double Lp;   // plan length
            public double Ext;  // extension amount
            public double RadOut;      // OUTSIDE radius (feet)
            public double RadVertical; // r_out * sqrt(1 - uz^2)
            public ElementId ElemId;
        }
        private class VertRec
        {
            public double X, Y;
            public XYZ C0, C1;
            public ElementId ElemId;
        }

        // ===== helpers =====

        private static string Csv(string s)
        {
            if (s == null) s = "";
            if (s.Contains(",") || s.Contains("\"") || s.Contains("\n"))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        // Build undirected adjacency across all fabrication parts
        private static Dictionary<ElementId, HashSet<ElementId>> BuildAdjacency(List<FabricationPart> parts)
        {
            var adj = new Dictionary<ElementId, HashSet<ElementId>>();
            foreach (var fp in parts)
            {
                var id = fp.Id;
                if (!adj.ContainsKey(id)) adj[id] = new HashSet<ElementId>(new ElemIdComparer());

                ConnectorSet cons = null;
                try { cons = fp.ConnectorManager?.Connectors; } catch { cons = null; }
                if (cons == null) continue;

                foreach (Connector c in cons)
                {
                    if (c == null) continue;
                    ConnectorSet refs = null;
                    try { refs = c.AllRefs; } catch { refs = null; }
                    if (refs == null) continue;

                    foreach (Connector r in refs)
                    {
                        var owner = r?.Owner as FabricationPart;
                        if (owner == null) continue;

                        var jd = owner.Id;
                        if (jd.IntegerValue == id.IntegerValue) continue;

                        if (!adj.ContainsKey(jd)) adj[jd] = new HashSet<ElementId>(new ElemIdComparer());
                        adj[id].Add(jd);
                        adj[jd].Add(id);
                    }
                }
            }
            return adj;
        }

        private class ElemIdComparer : IEqualityComparer<ElementId>
        {
            public bool Equals(ElementId x, ElementId y) => x.IntegerValue == y.IntegerValue;
            public int GetHashCode(ElementId obj) => obj.IntegerValue;
        }

        private static bool AreConnected(Dictionary<ElementId, HashSet<ElementId>> adj,
                                         ElementId a, ElementId b, int maxHops)
        {
            if (a.IntegerValue == b.IntegerValue) return true;
            if (!adj.TryGetValue(a, out var _)) return false;
            if (!adj.TryGetValue(b, out var _)) return false;

            var q = new Queue<(ElementId id, int d)>();
            var seen = new HashSet<ElementId>(new ElemIdComparer());
            q.Enqueue((a, 0));
            seen.Add(a);

            while (q.Count > 0)
            {
                var (cur, d) = q.Dequeue();
                if (!adj.TryGetValue(cur, out var nbrs)) continue;

                foreach (var n in nbrs)
                {
                    if (seen.Contains(n)) continue;
                    if (n.IntegerValue == b.IntegerValue) return true;
                    if (d + 1 < maxHops)
                    {
                        seen.Add(n);
                        q.Enqueue((n, d + 1));
                    }
                }
            }
            return false;
        }

        // Return connected straights reachable from a start part (by id)
        private static List<PipeRec> FindConnectedStraights(ElementId startId,
            Dictionary<ElementId, HashSet<ElementId>> adj,
            Dictionary<ElementId, PipeRec> straightById,
            int maxHops)
        {
            var result = new List<PipeRec>();
            if (!adj.TryGetValue(startId, out var _)) return result;

            var q = new Queue<(ElementId id, int d)>();
            var seen = new HashSet<ElementId>(new ElemIdComparer());
            q.Enqueue((startId, 0));
            seen.Add(startId);

            while (q.Count > 0)
            {
                var (cur, d) = q.Dequeue();
                if (!adj.TryGetValue(cur, out var nbrs)) continue;

                foreach (var n in nbrs)
                {
                    if (seen.Contains(n)) continue;
                    seen.Add(n);

                    if (straightById.TryGetValue(n, out var pr))
                        result.Add(pr);

                    if (d + 1 < maxHops)
                        q.Enqueue((n, d + 1));
                }
            }
            return result;
        }

        // STRAIGHT detector:
        // - must have exactly two connectors
        // - connector axes must be colinear & opposite (dot ≈ -1)
        private static bool IsStraight(FabricationPart fp)
        {
            if (fp == null) return false;

            IList<Connector> cons;
            try
            {
                var cm = fp.ConnectorManager?.Connectors;
                if (cm == null || cm.Size != 2) return false;
                cons = new List<Connector>();
                foreach (Connector c in cm) if (c != null) cons.Add(c);
                if (cons.Count != 2) return false;
            }
            catch { return false; }

            // Try axis test; if unavailable, fall back to "assume straight"
            if (TryGetAxis(cons[0], out XYZ d0) && TryGetAxis(cons[1], out XYZ d1))
            {
                d0 = SafeNorm(d0);
                d1 = SafeNorm(d1);
                if (d0 == null || d1 == null) return true; // fallback
                double dot = d0.X * d1.X + d0.Y * d1.Y + d0.Z * d1.Z;
                return dot <= -0.999; // opposite directions
            }

            return true;
        }

        private static bool TryGetAxis(Connector c, out XYZ dir)
        {
            dir = null;
            try
            {
                var cs = c.CoordinateSystem;
                if (cs != null)
                {
                    dir = cs.BasisZ; // connector axis
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static XYZ SafeNorm(XYZ v)
        {
            if (v == null) return null;
            double L = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            if (L < EPS) return null;
            return new XYZ(v.X / L, v.Y / L, v.Z / L);
        }

        private static List<Connector> GetConnectors(FabricationPart fp)
        {
            var list = new List<Connector>();
            try
            {
                if (fp?.ConnectorManager?.Connectors != null)
                    foreach (Connector c in fp.ConnectorManager.Connectors)
                        if (c != null) list.Add(c);
            }
            catch { }
            return list;
        }

        // ---- OUTSIDE radius (feet) with robust fallbacks ----
        private static double GetOutsideRadiusFeet(FabricationPart fp, List<Connector> cs)
        {
            double rad = 0.0;

            // 1) Largest connector radius (often OD/2, but not always)
            try
            {
                foreach (var c in cs)
                    if (c != null && c.Radius > rad)
                        rad = c.Radius; // feet
            }
            catch { }

            // 2) Prefer explicit diameter-like parameters on the part
            string[] names =
            {
                "Outside Diameter", "OutsideDiameter", "Outer Diameter", "OuterDiameter",
                "OD", "Outside Dia", "Outside dia", "Diameter", "Primary Diameter", "Nominal Diameter"
            };

            foreach (var n in names)
            {
                var p = fp.LookupParameter(n);
                if (p == null) continue;

                double odFt = 0.0;

                if (p.StorageType == StorageType.Double)
                {
                    // Revit stores doubles in internal feet
                    odFt = p.AsDouble();
                }
                else
                {
                    try
                    {
                        string vs = p.AsValueString();
                        if (TryParseLengthToFeet(vs, out double ft))
                            odFt = ft;
                    }
                    catch { }
                }

                if (odFt > 0)
                {
                    rad = Math.Max(rad, odFt * 0.5);
                    break;
                }
            }

            // 3) Last resort: a small, safe default (2" OD)
            if (rad <= 0) rad = (2.0 / 12.0) * 0.5;

            return rad;
        }

        private static bool TryParseLengthToFeet(string text, out double feet)
        {
            feet = 0.0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string s = text.Trim().ToLowerInvariant();

            var m = Regex.Match(s, @"([-+]?\d*\.?\d+)");
            if (!m.Success) return false;
            if (!double.TryParse(m.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
            {
                if (!double.TryParse(m.Value, out val)) return false;
            }

            if (s.Contains("mm"))
                feet = val / 304.8;      // mm → ft
            else if (s.Contains("cm"))
                feet = val / 30.48;      // cm → ft
            else if (s.Contains("m"))
                feet = val / 0.3048;     // m → ft
            else if (s.Contains("ft") || s.Contains("'"))
                feet = val;              // already feet
            else if (s.Contains("in") || s.Contains("\""))
                feet = val / 12.0;       // inches → ft
            else
                feet = (val < 48.0) ? val / 12.0 : val; // heuristic

            return true;
        }

        private static bool FarthestTwo(List<XYZ> pts, out XYZ a, out XYZ b)
        {
            a = b = null;
            double best = -1;
            for (int i = 0; i < pts.Count; ++i)
                for (int j = i + 1; j < pts.Count; ++j)
                {
                    double d = pts[i].DistanceTo(pts[j]);
                    if (d > best) { best = d; a = pts[i]; b = pts[j]; }
                }
            return best > 0;
        }

        private static XYZ Flatten(XYZ p) => new XYZ(p.X, p.Y, 0.0);
        private static double Dist2D(XYZ a, XYZ b) => Math.Sqrt((a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y));

        private static void Dir2D(XYZ a, XYZ b, out double ux, out double uy, out double len)
        {
            double vx = b.X - a.X, vy = b.Y - a.Y;
            len = Math.Sqrt(vx * vx + vy * vy);
            if (len < EPS) { ux = uy = 0; return; }
            ux = vx / len; uy = vy / len;
        }

        // rZ = r * sqrt(1 - uz^2), where uz is Z-component of the unit axis vector.
        private static double ComputeVerticalRadiusComponent(XYZ c0, XYZ c1, double rFeet)
        {
            double dx = c1.X - c0.X, dy = c1.Y - c0.Y, dz = c1.Z - c0.Z;
            double L3 = Math.Sqrt(dx * dx + dy * dy + dz * dz);
            if (L3 < EPS) return rFeet; // degenerate; treat as horizontal
            double uz = dz / L3; // [-1..1]
            double factor = Math.Sqrt(Math.Max(0.0, 1.0 - uz * uz)); // sin(angle to vertical)
            return rFeet * factor;
        }

        // segment–segment intersection (XY) on extended segments
        private static bool SegSegIntersect2D(XYZ A0, XYZ A1, XYZ B0, XYZ B1,
                                              out double ix, out double iy,
                                              out double ta, out double tb)
        {
            ix = iy = ta = tb = 0;
            double ax = A1.X - A0.X, ay = A1.Y - A0.Y;
            double bx = B1.X - B0.X, by = B1.Y - B0.Y;
            double det = ax * by - ay * bx;
            if (Math.Abs(det) < TOL) return false;
            double dx = B0.X - A0.X, dy = B0.Y - A0.Y;
            ta = (dx * by - dy * bx) / det;
            tb = (dx * ay - dy * ax) / det;
            if (ta < -TOL || ta > 1 + TOL || tb < -TOL || tb > 1 + TOL) return false;
            ix = A0.X + ta * ax;
            iy = A0.Y + ta * ay;
            return true;
        }

        // map param on extended line [0..1] to param on real 3D line (can be <0 or >1)
        private static double MapT_ExtToT3D(double tExt, double Lp, double extFt)
        {
            if (Lp < EPS) return 0.0;
            double s = tExt * (Lp + 2.0 * extFt) - extFt; // station along real span [0..Lp]
            return s / Lp; // fraction along run (same in 3D as in plan for a straight)
        }

        private static double BopZAtT(PipeRec pr, double t3)
        {
            double zc = pr.C0.Z + t3 * (pr.C1.Z - pr.C0.Z); // centerline Z at station
            return zc - pr.RadVertical; // slope-correct BOP using OUTSIDE radius
        }

        // distance from 2D point to 2D segment; returns param t on segment [0..1]
        private static double DistPtToSeg2D(double px, double py, XYZ A0, XYZ A1, out double t)
        {
            double vx = A1.X - A0.X, vy = A1.Y - A0.Y;
            double L2 = vx * vx + vy * vy;
            if (L2 < EPS) { t = 0.0; return Math.Sqrt((px - A0.X) * (px - A0.X) + (py - A0.Y) * (py - A0.Y)); }
            t = ((px - A0.X) * vx + (py - A0.Y) * vy) / L2;
            t = Math.Max(0.0, Math.Min(1.0, t));
            double qx = A0.X + t * vx, qy = A0.Y + t * vy;
            return Math.Sqrt((px - qx) * (px - qx) + (py - qy) * (py - qy));
        }

        private static bool NearAnyXY(List<XYZ> pts, XYZ pt, double tol)
        {
            foreach (var q in pts)
            {
                double dx = pt.X - q.X, dy = pt.Y - q.Y;
                if (Math.Sqrt(dx * dx + dy * dy) <= tol) return true;
            }
            return false;
        }

        // Merge essentially identical stake points; keep the lower Z.
        private static List<XYZ> MergeNearby(List<XYZ> pts, double xyTolFt, double zTolFt)
        {
            var result = new List<XYZ>();
            foreach (var p in pts)
            {
                int idx = -1;
                for (int i = 0; i < result.Count; i++)
                {
                    var q = result[i];
                    double dxy = Math.Sqrt((p.X - q.X) * (p.X - q.X) + (p.Y - q.Y) * (p.Y - q.Y));
                    if (dxy <= xyTolFt && Math.Abs(p.Z - q.Z) <= zTolFt)
                    {
                        idx = i;
                        break;
                    }
                }
                if (idx < 0)
                {
                    result.Add(p);
                }
                else
                {
                    if (p.Z < result[idx].Z)
                        result[idx] = new XYZ(result[idx].X, result[idx].Y, p.Z);
                }
            }
            return result;
        }
    }
}
