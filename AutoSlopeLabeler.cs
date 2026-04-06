using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;

// FIX 1: 'Surface' exists in both AutoCAD and Civil namespaces — alias Civil one
using CivilSurface = Autodesk.Civil.DatabaseServices.Surface;

[assembly: ExtensionApplication(typeof(AutoSlopeLabeler.AutoSlopeApp))]
[assembly: CommandClass(typeof(AutoSlopeLabeler.AutoSlopeCommands))]

namespace AutoSlopeLabeler
{
    // ─────────────────────────────────────────────────────────────
    // Plugin entry point
    // ─────────────────────────────────────────────────────────────
    public class AutoSlopeApp : IExtensionApplication
    {
        public void Initialize()
        {
            Application.DocumentManager.MdiActiveDocument?.Editor.WriteMessage(
                "\nAutoSlope Labeler loaded. Type AUTOSLOPE to run.\n");
        }

        public void Terminate() { }
    }

    // ─────────────────────────────────────────────────────────────
    // Main command
    // ─────────────────────────────────────────────────────────────
    public class AutoSlopeCommands
    {
        [CommandMethod("AUTOSLOPE")]
        public static void AutoSlope()
        {
            var doc   = Application.DocumentManager.MdiActiveDocument;
            var db    = doc.Database;
            var ed    = doc.Editor;
            var civil = CivilApplication.ActiveDocument;

            ed.WriteMessage("\n--- AutoSlope Labeler ---\n");

            // FIX: removed unused 'db' param from PickSurface and PickAlignment
            var surfaceId = PickSurface(ed);
            if (surfaceId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            var labelStyleId = PickSlopeLabelStyle(ed, db, civil);
            if (labelStyleId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            var centerlineId = PickAlignment(ed, "Select CENTERLINE alignment");
            if (centerlineId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            var lipLineId = PickAlignment(ed, "Select LIP LINE alignment (one side)");
            if (lipLineId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            double startStation, endStation, interval;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var cl     = (Alignment)tr.GetObject(centerlineId, OpenMode.ForRead);
                double clS = cl.StartingStation;
                double clE = cl.EndingStation;
                ed.WriteMessage($"\nCenterline runs from {clS:F3} to {clE:F3}.");
                tr.Commit();

                var r1 = ed.GetDouble(new PromptDoubleOptions(
                    $"\nEnter start station [{clS:F3}]")
                    { DefaultValue = clS, AllowNone = true });
                if (r1.Status == PromptStatus.Cancel) return;
                startStation = r1.Status == PromptStatus.None ? clS : r1.Value;

                var r2 = ed.GetDouble(new PromptDoubleOptions(
                    $"\nEnter end station [{clE:F3}]")
                    { DefaultValue = clE, AllowNone = true });
                if (r2.Status == PromptStatus.Cancel) return;
                endStation = r2.Status == PromptStatus.None ? clE : r2.Value;
            }

            var r3 = ed.GetDouble(new PromptDoubleOptions(
                "\nEnter label interval in meters [10]")
                { DefaultValue = 10.0, AllowNone = true });
            if (r3.Status == PromptStatus.Cancel) return;
            interval = r3.Status == PromptStatus.None ? 10.0 : r3.Value;

            if (interval <= 0)
            { ed.WriteMessage("\nInterval must be > 0.\n"); return; }
            if (endStation <= startStation)
            { ed.WriteMessage("\nEnd station must be > start station.\n"); return; }

            int count = PlaceLabels(db, ed, surfaceId, labelStyleId,
                                    centerlineId, lipLineId,
                                    startStation, endStation, interval);

            ed.WriteMessage($"\nDone. {count} slope label(s) placed.\n");
        }

        // ─────────────────────────────────────────────────────────
        // Core: step along centerline, find perpendicular lip point,
        //       place native Civil 3D two-point surface slope label
        // ─────────────────────────────────────────────────────────
        private static int PlaceLabels(
            Database db, Editor ed,
            ObjectId surfaceId, ObjectId labelStyleId,
            ObjectId centerlineId, ObjectId lipLineId,
            double startStation, double endStation, double interval)
        {
            int placed = 0;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var clAlignment  = (Alignment)tr.GetObject(centerlineId, OpenMode.ForRead);
                var lipAlignment = (Alignment)tr.GetObject(lipLineId,    OpenMode.ForRead);

                double currentStation = startStation;

                while (currentStation <= endStation + 1e-6)
                {
                    try
                    {
                        // FIX 2: PointLocation uses ref, not out
                        double clX = 0, clY = 0;
                        clAlignment.PointLocation(currentStation, 0.0, ref clX, ref clY);
                        var clPoint2d = new Point2d(clX, clY);

                        // FIX 3: No GetTangentDirection in Civil API —
                        //         compute from two nearby points instead
                        double tangentDir = GetAlignmentDirection(
                            clAlignment, currentStation);
                        double perpDir = tangentDir + Math.PI / 2.0;

                        Point2d lipPoint2d = FindPerpendicularIntersection(
                            clPoint2d, perpDir, lipAlignment);

                        if (lipPoint2d == Point2d.Origin)
                        {
                            perpDir    = tangentDir - Math.PI / 2.0;
                            lipPoint2d = FindPerpendicularIntersection(
                                clPoint2d, perpDir, lipAlignment);
                        }

                        if (lipPoint2d == Point2d.Origin)
                        {
                            ed.WriteMessage(
                                $"\n  [Skip] Station {currentStation:F3}: " +
                                "no lip line intersection found.");
                            currentStation += interval;
                            continue;
                        }

                        // Civil 3D 2026 SurfaceSlopeLabel.Create signature:
                        //   (ObjectId surfaceId, Point2d pt1, Point2d pt2, ObjectId labelStyleId)
                        ObjectId newLabelId = SurfaceSlopeLabel.Create(
                            surfaceId,
                            clPoint2d,
                            lipPoint2d,
                            labelStyleId);

                        if (newLabelId != ObjectId.Null)
                            placed++;
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage(
                            $"\n  [Warning] Station {currentStation:F3}: {ex.Message}");
                    }

                    currentStation += interval;
                }

                tr.Commit();
            }

            return placed;
        }

        // ─────────────────────────────────────────────────────────
        // FIX 3: Compute tangent direction via two closely-spaced
        //        PointLocation calls (finite difference approach).
        //        Civil 3D API has no single GetTangentDirection method.
        // ─────────────────────────────────────────────────────────
        private static double GetAlignmentDirection(
            Alignment al, double station)
        {
            double delta  = 0.01; // 1 cm
            double sStart = al.StartingStation;
            double sEnd   = al.EndingStation;

            double s1 = Math.Max(sStart, station - delta);
            double s2 = Math.Min(sEnd,   station + delta);

            double x1 = 0, y1 = 0, x2 = 0, y2 = 0;
            al.PointLocation(s1, 0.0, ref x1, ref y1);
            al.PointLocation(s2, 0.0, ref x2, ref y2);

            return Math.Atan2(y2 - y1, x2 - x1);
        }

        // ─────────────────────────────────────────────────────────
        // Coarse (1 m) then fine (0.1 m) search for the lip line
        // point that lies on the perpendicular ray from clPoint.
        // Returns Point2d.Origin if nothing found within 2 m of ray.
        // ─────────────────────────────────────────────────────────
        private static Point2d FindPerpendicularIntersection(
            Point2d clPoint, double perpDirection,
            Alignment lipAlignment)
        {
            var rayDir = new Vector2d(
                Math.Cos(perpDirection),
                Math.Sin(perpDirection));

            double lipStart    = lipAlignment.StartingStation;
            double lipEnd      = lipAlignment.EndingStation;
            double bestStation = lipStart;
            double minDist     = double.MaxValue;

            // Coarse pass — 1 m steps
            for (double s = lipStart; s <= lipEnd; s += 1.0)
            {
                double lx = 0, ly = 0;
                try { lipAlignment.PointLocation(s, 0.0, ref lx, ref ly); }
                catch { continue; }

                var    diff     = new Point2d(lx, ly) - clPoint;
                double t        = diff.DotProduct(rayDir);
                double perpDist = Math.Abs(diff.X * rayDir.Y - diff.Y * rayDir.X);

                if (t > 0 && perpDist < minDist)
                {
                    minDist     = perpDist;
                    bestStation = s;
                }
            }

            if (minDist > 2.0)
                return Point2d.Origin;

            // Fine pass — 0.1 m steps around best candidate
            double fineStart      = Math.Max(lipStart, bestStation - 1.0);
            double fineEnd        = Math.Min(lipEnd,   bestStation + 1.0);
            double refinedStation = bestStation;
            minDist = double.MaxValue;

            for (double s = fineStart; s <= fineEnd; s += 0.1)
            {
                double lx = 0, ly = 0;
                try { lipAlignment.PointLocation(s, 0.0, ref lx, ref ly); }
                catch { continue; }

                var    diff     = new Point2d(lx, ly) - clPoint;
                double t        = diff.DotProduct(rayDir);
                double perpDist = Math.Abs(diff.X * rayDir.Y - diff.Y * rayDir.X);

                if (t > 0 && perpDist < minDist)
                {
                    minDist         = perpDist;
                    refinedStation  = s;
                }
            }

            double rx = 0, ry = 0;
            try { lipAlignment.PointLocation(refinedStation, 0.0, ref rx, ref ry); }
            catch { return Point2d.Origin; }

            return new Point2d(rx, ry);
        }

        // ─────────────────────────────────────────────────────────
        // UI helpers — static, no instance data needed
        // ─────────────────────────────────────────────────────────

        private static ObjectId PickSurface(Editor ed)
        {
            var opts = new PromptEntityOptions("\nSelect surface: ");
            opts.SetRejectMessage("\nNot a surface — try again.");
            opts.AddAllowedClass(typeof(TinSurface),  exactMatch: false);
            opts.AddAllowedClass(typeof(GridSurface), exactMatch: false);

            var res = ed.GetEntity(opts);
            return res.Status == PromptStatus.OK ? res.ObjectId : ObjectId.Null;
        }

        private static ObjectId PickAlignment(Editor ed, string prompt)
        {
            var opts = new PromptEntityOptions($"\n{prompt}: ");
            opts.SetRejectMessage("\nNot an alignment — try again.");
            opts.AddAllowedClass(typeof(Alignment), exactMatch: false);

            var res = ed.GetEntity(opts);
            return res.Status == PromptStatus.OK ? res.ObjectId : ObjectId.Null;
        }

        private static ObjectId PickSlopeLabelStyle(
            Editor ed, Database db, CivilDocument civil)
        {
            var styleNames = new List<string>();
            var styleIds   = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var slopeStyles =
                    civil.Styles.LabelStyles.SurfaceLabelStyles.SlopeLabelStyles;

                foreach (ObjectId id in slopeStyles)
                {
                    // FIX 5: use pattern matching 'is' — cleaner, no null warning
                    if (tr.GetObject(id, OpenMode.ForRead) is LabelStyle style)
                    {
                        styleNames.Add(style.Name);
                        styleIds.Add(id);
                    }
                }
                tr.Commit();
            }

            if (styleNames.Count == 0)
            {
                ed.WriteMessage(
                    "\nNo surface slope label styles found in this drawing." +
                    "\nGo to: Toolspace → Settings → Surface → Label Styles → Slope" +
                    "\nRight-click → New, create a style, then run AUTOSLOPE again.");
                return ObjectId.Null;
            }

            ed.WriteMessage("\nAvailable slope label styles:");
            for (int i = 0; i < styleNames.Count; i++)
                ed.WriteMessage($"\n  [{i + 1}] {styleNames[i]}");

            var res = ed.GetInteger(new PromptIntegerOptions(
                $"\nEnter number (1-{styleNames.Count})")
            {
                LowerLimit   = 1,
                UpperLimit   = styleNames.Count,
                DefaultValue = 1,
                AllowNone    = true
            });

            if (res.Status == PromptStatus.Cancel) return ObjectId.Null;
            int choice = res.Status == PromptStatus.None ? 1 : res.Value;

            ed.WriteMessage($"\nUsing style: {styleNames[choice - 1]}");
            return styleIds[choice - 1];
        }
    }
}
