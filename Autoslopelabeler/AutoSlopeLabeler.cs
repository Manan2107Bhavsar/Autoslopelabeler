using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.DatabaseServices.Styles;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;

using CivilSurface = Autodesk.Civil.DatabaseServices.Surface;

[assembly: ExtensionApplication(typeof(AutoSlopeLabeler.AutoSlopeApp))]
[assembly: CommandClass(typeof(AutoSlopeLabeler.AutoSlopeCommands))]

namespace AutoSlopeLabeler
{
    // ─────────────────────────────────────────────────────────────
    // Small data class — one row in the surface picker list
    // ─────────────────────────────────────────────────────────────
    internal class SurfaceItem
    {
        public string Name { get; }
        public string TypeLabel { get; }   // "TIN" or "Grid"
        public ObjectId Id { get; }

        public SurfaceItem(string name, string typeLabel, ObjectId id)
        {
            Name = name;
            TypeLabel = typeLabel;
            Id = id;
        }

        // What the ListBox shows
        public override string ToString() => $"{Name}  [{TypeLabel}]";
    }

    // ─────────────────────────────────────────────────────────────
    // WPF dialog — built entirely in code, no XAML file needed.
    // Shows all surfaces in the drawing; user clicks one and OK.
    // ─────────────────────────────────────────────────────────────
    internal class SurfacePickerDialog : Window
    {
        private readonly ListBox _list;
        public ObjectId SelectedSurfaceId { get; private set; } = ObjectId.Null;

        public SurfacePickerDialog(List<SurfaceItem> surfaces)
        {
            // ── Window chrome ─────────────────────────────────────
            Title = "Select Surface";
            Width = 380;
            Height = 320;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(245, 245, 245));

            // ── Layout ────────────────────────────────────────────
            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = root;

            // ── Surface list ──────────────────────────────────────
            var border = new Border
            {
                Margin = new Thickness(12, 12, 12, 6),
                BorderBrush = new SolidColorBrush(Color.FromRgb(180, 180, 180)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Background = Brushes.White
            };
            Grid.SetRow(border, 0);
            root.Children.Add(border);

            _list = new ListBox
            {
                FontSize = 13,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Padding = new Thickness(4)
            };
            foreach (var s in surfaces)
                _list.Items.Add(s);

            // Select first item by default
            if (_list.Items.Count > 0)
                _list.SelectedIndex = 0;

            // Double-click = instant confirm
            _list.MouseDoubleClick += (_, __) => Confirm();
            border.Child = _list;

            // ── Button row ────────────────────────────────────────
            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(12, 4, 12, 12)
            };
            Grid.SetRow(buttonPanel, 1);
            root.Children.Add(buttonPanel);

            var btnOk = new Button
            {
                Content = "OK",
                Width = 80,
                Height = 28,
                Margin = new Thickness(0, 0, 8, 0),
                IsDefault = true
            };
            btnOk.Click += (_, __) => Confirm();
            buttonPanel.Children.Add(btnOk);

            var btnCancel = new Button
            {
                Content = "Cancel",
                Width = 80,
                Height = 28,
                IsCancel = true
            };
            btnCancel.Click += (_, __) => { DialogResult = false; Close(); };
            buttonPanel.Children.Add(btnCancel);
        }

        private void Confirm()
        {
            if (_list.SelectedItem is SurfaceItem item)
            {
                SelectedSurfaceId = item.Id;
                DialogResult = true;
                Close();
            }
            else
            {
                MessageBox.Show(
                    "Please select a surface from the list.",
                    "AutoSlope",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }
    }

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
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;
            var civil = CivilApplication.ActiveDocument;

            ed.WriteMessage("\n--- AutoSlope Labeler ---\n");

            // ── 1. Surface picker — WPF dialog listing all surfaces ──
            var surfaceId = PickSurfaceFromDialog(db);
            if (surfaceId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            // ── 2. Slope label style ──────────────────────────────────
            var labelStyleId = PickSlopeLabelStyle(ed, db, civil);
            if (labelStyleId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            // ── 3. Centerline alignment ───────────────────────────────
            var centerlineId = PickAlignment(ed, "Select CENTERLINE alignment");
            if (centerlineId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            // ── 4. Lip line alignment ─────────────────────────────────
            var lipLineId = PickAlignment(ed, "Select LIP LINE alignment (one side)");
            if (lipLineId == ObjectId.Null) { ed.WriteMessage("\nCancelled.\n"); return; }

            // ── 5. Station range and interval ─────────────────────────
            double startStation, endStation, interval;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var cl = (Alignment)tr.GetObject(centerlineId, OpenMode.ForRead);
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

            // ── 6. Place labels ───────────────────────────────────────
            int count = PlaceLabels(db, ed, surfaceId, labelStyleId,
                                    centerlineId, lipLineId,
                                    startStation, endStation, interval);

            ed.WriteMessage($"\nDone. {count} slope label(s) placed.\n");
        }

        // ─────────────────────────────────────────────────────────
        // Opens WPF dialog populated with every TIN/Grid surface
        // found in the drawing.  Returns ObjectId.Null on cancel.
        // Surface visibility does NOT matter — all are listed.
        // ─────────────────────────────────────────────────────────
        private static ObjectId PickSurfaceFromDialog(Database db)
        {
            var items = new List<SurfaceItem>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Walk the entire model-space block to find surfaces
                var bt = (BlockTable)tr.GetObject(
                    db.BlockTableId, OpenMode.ForRead);
                var btr = (BlockTableRecord)tr.GetObject(
                    bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId id in btr)
                {
                    var obj = tr.GetObject(id, OpenMode.ForRead);

                    if (obj is TinSurface tin)
                        items.Add(new SurfaceItem(tin.Name, "TIN", id));
                    else if (obj is GridSurface grid)
                        items.Add(new SurfaceItem(grid.Name, "Grid", id));
                }

                tr.Commit();
            }

            if (items.Count == 0)
            {
                MessageBox.Show(
                    "No surfaces found in this drawing.",
                    "AutoSlope",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return ObjectId.Null;
            }

            // Show dialog on the AutoCAD UI thread
            var dlg = new SurfacePickerDialog(items);
            bool? result = dlg.ShowDialog();

            return result == true ? dlg.SelectedSurfaceId : ObjectId.Null;
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
                var clAlignment = (Alignment)tr.GetObject(centerlineId, OpenMode.ForRead);
                var lipAlignment = (Alignment)tr.GetObject(lipLineId, OpenMode.ForRead);

                double currentStation = startStation;

                while (currentStation <= endStation + 1e-6)
                {
                    try
                    {
                        double clX = 0, clY = 0;
                        clAlignment.PointLocation(currentStation, 0.0, ref clX, ref clY);
                        var clPoint2d = new Point2d(clX, clY);

                        double tangentDir = GetAlignmentDirection(clAlignment, currentStation);
                        double perpDir = tangentDir + Math.PI / 2.0;

                        Point2d lipPoint2d = FindPerpendicularIntersection(
                            clPoint2d, perpDir, lipAlignment);

                        if (lipPoint2d == Point2d.Origin)
                        {
                            perpDir = tangentDir - Math.PI / 2.0;
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
        // Tangent direction via finite difference (Civil API has no
        // single GetTangentDirection method)
        // ─────────────────────────────────────────────────────────
        private static double GetAlignmentDirection(Alignment al, double station)
        {
            double delta = 0.01;
            double s1 = Math.Max(al.StartingStation, station - delta);
            double s2 = Math.Min(al.EndingStation, station + delta);
            double x1 = 0, y1 = 0, x2 = 0, y2 = 0;
            al.PointLocation(s1, 0.0, ref x1, ref y1);
            al.PointLocation(s2, 0.0, ref x2, ref y2);
            return Math.Atan2(y2 - y1, x2 - x1);
        }

        // ─────────────────────────────────────────────────────────
        // Coarse (1 m) + fine (0.1 m) search for the lip line point
        // on the perpendicular ray from clPoint.
        // Returns Point2d.Origin if nothing found within 2 m of ray.
        // ─────────────────────────────────────────────────────────
        private static Point2d FindPerpendicularIntersection(
            Point2d clPoint, double perpDirection, Alignment lipAlignment)
        {
            var rayDir = new Vector2d(Math.Cos(perpDirection), Math.Sin(perpDirection));
            double lipStart = lipAlignment.StartingStation;
            double lipEnd = lipAlignment.EndingStation;
            double bestStn = lipStart;
            double minDist = double.MaxValue;

            for (double s = lipStart; s <= lipEnd; s += 1.0)
            {
                double lx = 0, ly = 0;
                try { lipAlignment.PointLocation(s, 0.0, ref lx, ref ly); }
                catch { continue; }

                var diff = new Point2d(lx, ly) - clPoint;
                double t = diff.DotProduct(rayDir);
                double perpDist = Math.Abs(diff.X * rayDir.Y - diff.Y * rayDir.X);

                if (t > 0 && perpDist < minDist) { minDist = perpDist; bestStn = s; }
            }

            if (minDist > 2.0) return Point2d.Origin;

            double fineS = Math.Max(lipStart, bestStn - 1.0);
            double fineE = Math.Min(lipEnd, bestStn + 1.0);
            double refStn = bestStn;
            minDist = double.MaxValue;

            for (double s = fineS; s <= fineE; s += 0.1)
            {
                double lx = 0, ly = 0;
                try { lipAlignment.PointLocation(s, 0.0, ref lx, ref ly); }
                catch { continue; }

                var diff = new Point2d(lx, ly) - clPoint;
                double t = diff.DotProduct(rayDir);
                double perpDist = Math.Abs(diff.X * rayDir.Y - diff.Y * rayDir.X);

                if (t > 0 && perpDist < minDist) { minDist = perpDist; refStn = s; }
            }

            double rx = 0, ry = 0;
            try { lipAlignment.PointLocation(refStn, 0.0, ref rx, ref ry); }
            catch { return Point2d.Origin; }

            return new Point2d(rx, ry);
        }

        // ─────────────────────────────────────────────────────────
        // Alignment picker — viewport click (no change needed here,
        // alignments are always visible as entities)
        // ─────────────────────────────────────────────────────────
        private static ObjectId PickAlignment(Editor ed, string prompt)
        {
            var opts = new PromptEntityOptions($"\n{prompt}: ");
            opts.SetRejectMessage("\nNot an alignment — try again.");
            opts.AddAllowedClass(typeof(Alignment), exactMatch: false);

            var res = ed.GetEntity(opts);
            return res.Status == PromptStatus.OK ? res.ObjectId : ObjectId.Null;
        }

        // ─────────────────────────────────────────────────────────
        // Label style picker — command line list (unchanged)
        // ─────────────────────────────────────────────────────────
        private static ObjectId PickSlopeLabelStyle(
            Editor ed, Database db, CivilDocument civil)
        {
            var styleNames = new List<string>();
            var styleIds = new List<ObjectId>();

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var slopeStyles =
                    civil.Styles.LabelStyles.SurfaceLabelStyles.SlopeLabelStyles;

                foreach (ObjectId id in slopeStyles)
                {
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

            // Auto-select if only one style exists — skip the prompt entirely
            if (styleNames.Count == 1)
            {
                ed.WriteMessage($"\nUsing style: {styleNames[0]}");
                return styleIds[0];
            }

            ed.WriteMessage("\nAvailable slope label styles:");
            for (int i = 0; i < styleNames.Count; i++)
                ed.WriteMessage($"\n  [{i + 1}] {styleNames[i]}");

            var res = ed.GetInteger(new PromptIntegerOptions(
                $"\nEnter number (1-{styleNames.Count})")
            {
                LowerLimit = 1,
                UpperLimit = styleNames.Count,
                DefaultValue = 1,
                AllowNone = true
            });

            if (res.Status == PromptStatus.Cancel) return ObjectId.Null;
            int choice = res.Status == PromptStatus.None ? 1 : res.Value;

            ed.WriteMessage($"\nUsing style: {styleNames[choice - 1]}");
            return styleIds[choice - 1];
        }
    }
}