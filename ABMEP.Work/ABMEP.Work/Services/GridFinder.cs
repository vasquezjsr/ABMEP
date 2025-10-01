// File: ABMEP.Work/Services/GridFinder.cs
// Target: .NET Framework 4.8
// Revit: RevitAPI.dll
//
// Purpose:
//  - Utilities to find the nearest Grid(s) to a given point (e.g., sleeve center).
//  - Supports "nearest overall" as well as "nearest orthogonal X/Y" grids.
//  - Classifies grids by direction (X-like vs Y-like) using a tolerant angle test.
//  - All math is 2D (XY plane); Z is ignored for distance/classification purposes.
//
// Usage sketch (later step):
//   var finder = new GridFinder(doc);
//   var pt = sleeveLocationPoint; // XYZ
//   var both = finder.FindNearestOrthogonalGrids(pt);
//   var nearestX = both.Item1;
//   var nearestY = both.Item2;
//
//   // Then you’ll build References and create Dimensions in your command.

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace ABMEP.Work.Services
{
    public class GridFinder
    {
        private readonly Document _doc;
        private readonly List<Grid> _grids; // cached

        /// <summary>
        /// Angle (degrees) within which a grid is considered “X-like” or “Y-like”.
        /// 0° ~ X direction, 90° ~ Y direction. We allow 30° tolerance.
        /// </summary>
        private const double AxisToleranceDeg = 30.0;

        public GridFinder(Document doc)
        {
            _doc = doc;
            _grids = new FilteredElementCollector(_doc)
                .OfClass(typeof(Grid))
                .Cast<Grid>()
                .Where(g => g != null && g.Curve != null)
                .ToList();
        }

        /// <summary>
        /// Returns all grids (cached).
        /// </summary>
        public IList<Grid> AllGrids
        {
            get { return _grids; }
        }

        /// <summary>
        /// Find the single nearest grid (by shortest distance from point to grid curve, XY only).
        /// </summary>
        public Grid FindNearestGrid(XYZ point)
        {
            if (_grids.Count == 0) return null;

            Grid best = null;
            double bestDist = double.MaxValue;

            foreach (var g in _grids)
            {
                double d = DistanceXY(g.Curve, point);
                if (d < bestDist)
                {
                    bestDist = d;
                    best = g;
                }
            }
            return best;
        }

        /// <summary>
        /// Find the nearest “X-like” and “Y-like” grids to the given point (XY only).
        /// Returns (nearestX, nearestY). Either item may be null if no suitable grid exists.
        /// </summary>
        public Tuple<Grid, Grid> FindNearestOrthogonalGrids(XYZ point)
        {
            Grid bestX = null;
            Grid bestY = null;
            double bestXDist = double.MaxValue;
            double bestYDist = double.MaxValue;

            foreach (var g in _grids)
            {
                var kind = ClassifyGridDirection(g);
                if (kind == GridAxisKind.Unknown) continue;

                double d = DistanceXY(g.Curve, point);

                if (kind == GridAxisKind.XLike)
                {
                    if (d < bestXDist)
                    {
                        bestXDist = d;
                        bestX = g;
                    }
                }
                else if (kind == GridAxisKind.YLike)
                {
                    if (d < bestYDist)
                    {
                        bestYDist = d;
                        bestY = g;
                    }
                }
            }

            return Tuple.Create(bestX, bestY);
        }

        /// <summary>
        /// Find the nearest grid constrained to “X-like” or “Y-like”.
        /// </summary>
        public Grid FindNearestGridAlongAxis(XYZ point, GridAxisKind axisKind)
        {
            if (axisKind != GridAxisKind.XLike && axisKind != GridAxisKind.YLike) return null;

            Grid best = null;
            double bestDist = double.MaxValue;

            foreach (var g in _grids)
            {
                if (ClassifyGridDirection(g) != axisKind) continue;

                double d = DistanceXY(g.Curve, point);
                if (d < bestDist)
                {
                    bestDist = d;
                    best = g;
                }
            }
            return best;
        }

        /// <summary>
        /// Classify a grid as X-like (parallel to model X), Y-like (parallel to model Y), or Unknown.
        /// For arcs/curved grids, returns Unknown.
        /// </summary>
        public GridAxisKind ClassifyGridDirection(Grid grid)
        {
            if (grid == null || grid.Curve == null) return GridAxisKind.Unknown;
            var c = grid.Curve;

            // Only handle linear grids for axis classification. Curved grids => Unknown.
            var line = c as Line;
            if (line == null)
                return GridAxisKind.Unknown;

            // Direction in XY plane, normalized
            var dir = (line.Direction ?? XYZ.BasisX).Normalize();
            var dir2D = new XYZ(dir.X, dir.Y, 0.0).Normalize();
            if (dir2D.IsZeroLength()) return GridAxisKind.Unknown;

            // Compare with world X and Y
            double angToX = AngleDegrees(dir2D, new XYZ(1, 0, 0));
            double angToY = AngleDegrees(dir2D, new XYZ(0, 1, 0));

            // Bring angles into [0,90]
            angToX = Acute(angToX);
            angToY = Acute(angToY);

            if (angToX <= AxisToleranceDeg) return GridAxisKind.XLike;
            if (angToY <= AxisToleranceDeg) return GridAxisKind.YLike;

            return GridAxisKind.Unknown;
        }

        // ---------------- internal math helpers ----------------

        private static double DistanceXY(Curve curve, XYZ p)
        {
            if (curve == null || p == null) return double.MaxValue;

            // Project onto XY plane and measure in 2D
            var p2 = new XYZ(p.X, p.Y, 0.0);

            // Param on curve closest to p in 3D, then flatten to XY
            // For straight lines, this is fine; for arcs it’s still valid distance along the arc in 3D.
            try
            {
                var proj = curve.Project(p);
                if (proj != null)
                {
                    var closest = proj.XYZPoint;
                    var c2 = new XYZ(closest.X, closest.Y, 0.0);
                    return c2.DistanceTo(p2);
                }
            }
            catch { /* fall through */ }

            // Fallback: use endpoints (rough)
            try
            {
                var e0 = curve.GetEndPoint(0);
                var e1 = curve.GetEndPoint(1);
                var e0_2 = new XYZ(e0.X, e0.Y, 0.0);
                var e1_2 = new XYZ(e1.X, e1.Y, 0.0);
                return Math.Min(e0_2.DistanceTo(p2), e1_2.DistanceTo(p2));
            }
            catch { }

            return double.MaxValue;
        }

        private static double AngleDegrees(XYZ a, XYZ b)
        {
            double dot = a.X * b.X + a.Y * b.Y + a.Z * b.Z;
            double la = Math.Sqrt(a.X * a.X + a.Y * a.Y + a.Z * a.Z);
            double lb = Math.Sqrt(b.X * b.X + b.Y * b.Y + b.Z * b.Z);
            if (la <= 1e-9 || lb <= 1e-9) return 0.0;

            double c = dot / (la * lb);
            if (c > 1.0) c = 1.0;
            if (c < -1.0) c = -1.0;
            return RadToDeg(Math.Acos(c));
        }

        private static double RadToDeg(double r) { return r * (180.0 / Math.PI); }

        private static double Acute(double angDeg)
        {
            // Normalize angle to [0, 180], then map to [0, 90]
            double a = angDeg % 180.0;
            if (a < 0) a += 180.0;
            return (a > 90.0) ? (180.0 - a) : a;
        }
    }

    /// <summary>
    /// Simple axis classification for grids.
    /// </summary>
    public enum GridAxisKind
    {
        Unknown = 0,
        XLike = 1,
        YLike = 2
    }

    internal static class XyzExtensions
    {
        public static bool IsZeroLength(this XYZ v)
        {
            return (Math.Abs(v.X) < 1e-10 && Math.Abs(v.Y) < 1e-10 && Math.Abs(v.Z) < 1e-10);
        }
    }
}
