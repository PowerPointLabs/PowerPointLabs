using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.DataSources;
using PowerPointLabs.Models;
using PPExtraEventHelper;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Graphics = PowerPointLabs.Utils.Graphics;

namespace PowerPointLabs.DrawingsLab
{
    internal class DrawingsLabMain
    {
        private static Dictionary<Native.VirtualKey, ControlGroup> ControlGroups = new Dictionary<Native.VirtualKey, ControlGroup>();  

        private struct ControlGroup
        {
            public readonly int SlideId;
            public readonly HashSet<int> ShapeIds;

            public ControlGroup(int slideId, HashSet<int> shapeIds)
            {
                SlideId = slideId;
                ShapeIds = shapeIds;
            }
        }


        public static DrawingsLabDataSource DataSource
        {
            get { return DrawingsPaneWPF.dataSource; }
        }

        #region API

        public static void TestControlId()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;
            var cmd = Graphics.GetText(selection.ShapeRange[1]);
            //commandBars.ExecuteMso("MakeSegmentCurved");
            //commandBars.ExecuteMso("ShapeStraightConnector");
            try
            {
                Debug.WriteLine("Execute: " + cmd);
                commandBars.ExecuteMso(cmd);
            }
            catch (Exception e)
            {
                Debug.WriteLine("ERROR!");
                Debug.WriteLine(e);
            }
        }

        public static void SwitchToLineTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeStraightConnector");
        }

        public static void SwitchToArrowTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeStraightConnectorArrow");
        }

        public static void SwitchToRectangleTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeRectangle");
        }

        public static void SwitchToRoundedRectangleTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeRoundedRectangle");
        }

        public static void SwitchToCircleTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeOval");
        }

        public static void SwitchToTextboxTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("TextBoxInsertHorizontal");
        }

        public static void AddText()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText) return;

            var text = DrawingsLabDialogs.ShowInsertTextDialog();
            if (text == null) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in selection.ShapeRange.Cast<Shape>())
            {
                try
                {
                    Graphics.SetText(shape, text);
                }
                catch (ArgumentException e)
                {
                    Debug.WriteLine("Unable to write text to " + shape.Name);
                }
            }
        }

        public static void AddMath()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionText)
            {
                if (selection.Type != PpSelectionType.ppSelectionShapes)
                {
                    Error(TextCollection.DrawingsLabSelectExactlyOneShape);
                    return;
                }
                var shapes = selection.ShapeRange.Cast<Shape>().ToList();
                if (shapes.Count != 1)
                {
                    Error(TextCollection.DrawingsLabSelectExactlyOneShape);
                    return;
                }
            }

            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                var commandBars = Globals.ThisAddIn.Application.CommandBars;
                commandBars.ExecuteMso("EquationInsertNew");
            }
            catch (COMException e)
            {
                // Do nothing. EquationInsertNew throws an exception even as it succeeds.
            }
        }

        public static void RemoveText()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in selection.ShapeRange.Cast<Shape>())
            {
                Graphics.SetText(shape, String.Empty);
            }
        }

        public static void GroupShapes()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var slide = PowerPointCurrentPresentationInfo.CurrentSlide;
            var shapes = selection.ShapeRange.Cast<Shape>().ToList();

            if (shapes.Count < 2)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                slide.GroupShapes(shapes);
            }
            catch (UnauthorizedAccessException e)
            {
                Error(TextCollection.DrawingsLabErrorCannotGroup);
            }
        }

        public static void UngroupShapes()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            bool didSomething = false;
            foreach (var shape in selection.ShapeRange.Cast<Shape>().Where(shape => Graphics.IsAGroup(shape)))
            {
                shape.Ungroup();
                didSomething = true;
            }
            if (!didSomething)
            {
                Error(TextCollection.DrawingsLabErrorNothingUngrouped);
            }
        }


        public static void HideTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in selection.ShapeRange.Cast<Shape>())
            {
                shape.Visible = MsoTriState.msoFalse;
            }
        }

        public static void ShowAllTool()
        {
            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes.Cast<Shape>())
            {
                shape.Visible = MsoTriState.msoTrue;
            }
        }

        public static void CloneTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            PowerPointCurrentPresentationInfo.CurrentSlide.CopyShapesToSlide(selection.ShapeRange);
        }

        public static void MultiCloneExtendTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;
            var shapeList = selection.ShapeRange.Cast<Shape>().ToList();

            if (shapeList.Count % 2 != 0)
            {
                Error(TextCollection.DrawingsLabSelectTwoSetsOfShapes);
                return;
            }

            int clones = DrawingsLabDialogs.ShowMultiCloneNumericDialog();
            if (clones <= 0) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            int midpoint = shapeList.Count / 2;
            for (int i = 0; i < shapeList.Count / 2; ++i)
            {
                // Do the cloning for every pair of shapes (i, midpoint+i)
                var firstShape = shapeList[i];
                var secondShape = shapeList[midpoint + i];

                var newlyAddedShapes = new List<Shape>();
                for (int j = 0; j < clones; ++j)
                {
                    var newShape = firstShape.Duplicate()[1];
                    int index = j + 1;

                    newShape.Left = secondShape.Left + (secondShape.Left - firstShape.Left) * index;
                    newShape.Top = secondShape.Top + (secondShape.Top - firstShape.Top) * index;
                    newShape.Rotation = secondShape.Rotation + (secondShape.Rotation - firstShape.Rotation) * index;
                    newlyAddedShapes.Add(newShape);
                }

                // Set Z-Orders of newly created shapes.
                if (secondShape.ZOrderPosition < firstShape.ZOrderPosition)
                {
                    // first shape in front of last shape. Order the in-between shapes accordingly.
                    Shape prevShape = secondShape;
                    foreach (var shape in newlyAddedShapes)
                    {
                        Graphics.MoveZUntilBehind(shape, prevShape);
                        prevShape = shape;
                    }
                }
            }
        }

        public static void MultiCloneBetweenTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;
            var shapeList = selection.ShapeRange.Cast<Shape>().ToList();

            if (shapeList.Count % 2 != 0)
            {
                Error(TextCollection.DrawingsLabSelectTwoSetsOfShapes);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            int clones = DrawingsLabDialogs.ShowMultiCloneNumericDialog();
            if (clones <= 0) return;

            int divisions = clones + 1;

            int midpoint = shapeList.Count / 2;
            for (int i = 0; i < shapeList.Count / 2; ++i)
            {
                // Do the cloning for every pair of shapes (i, midpoint+i)
                var firstShape = shapeList[i];
                var lastShape = shapeList[midpoint + i];

                var newlyAddedShapes = new List<Shape>();
                for (int j = 0; j < clones; ++j)
                {
                    var newShape = firstShape.Duplicate()[1];
                    int index = j + 1;

                    newShape.Left = firstShape.Left + (lastShape.Left - firstShape.Left) / divisions * index;
                    newShape.Top = firstShape.Top + (lastShape.Top - firstShape.Top) / divisions * index;
                    newShape.Rotation = firstShape.Rotation + (lastShape.Rotation - firstShape.Rotation) / divisions * index;

                    newlyAddedShapes.Add(newShape);
                }

                // Set Z-Orders of newly created shapes.
                if (firstShape.ZOrderPosition < lastShape.ZOrderPosition)
                {
                    // last shape in front of first shape. Order the in-between shapes accordingly.
                    foreach (var shape in newlyAddedShapes)
                    {
                        Graphics.MoveZUntilBehind(shape, lastShape);
                    }
                }
                else
                {
                    // first shape in front of last shape. Order the in-between shapes accordingly.
                    newlyAddedShapes.Reverse();
                    foreach (var shape in newlyAddedShapes)
                    {
                        Graphics.MoveZUntilBehind(shape, firstShape);
                    }
                }
            }
        }

        public static void MultiCloneGridTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;
            var shapeList = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapeList.Count != 2)
            {
                Error(TextCollection.DrawingsLabSelectExactlyTwoShapes);
                return;
            }

            var sourceShape = shapeList[0];
            var targetShape = shapeList[1];

            var dialog = new MultiCloneGridDialog(sourceShape.Left, sourceShape.Top, targetShape.Left, targetShape.Top);
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            // Clone shapes in a grid.
            var newlyAddedShapes = new List<Shape>();

            var dx = targetShape.Left - sourceShape.Left;
            var dy = targetShape.Top - sourceShape.Top;
            int skipIndexX = 1;
            int skipIndexY = 1;
            if (!dialog.IsExtend)
            {
                // Is between.
                dx = dx/(dialog.XCopies - 1);
                dy = dy/(dialog.YCopies - 1);
                skipIndexX = dialog.XCopies - 1;
                skipIndexY = dialog.YCopies - 1;
            }

            for (int y = 0; y < dialog.YCopies; ++y)
            {
                for (int x = 0; x < dialog.XCopies; ++x)
                {
                    if (x == 0 && y == 0) continue;
                    if (x == skipIndexX && y == skipIndexY) continue;

                    var newShape = sourceShape.Duplicate()[1];
                    newShape.Left = sourceShape.Left + dx*x;
                    newShape.Top = sourceShape.Top + dy * y;
                    newlyAddedShapes.Add(newShape);

                }
            }

            // Set Z-Orders of newly created shapes.
            if (dialog.IsExtend)
            {
                // Multiclone Extend
                if (sourceShape.ZOrderPosition < targetShape.ZOrderPosition)
                {
                    if (newlyAddedShapes.Count >= dialog.YCopies)
                    {
                        for (int i = 0; i < dialog.YCopies; ++i)
                        {
                            Graphics.MoveZToJustBehind(newlyAddedShapes[i], targetShape);
                        }
                    }
                }
                else
                {
                    // first shape in front of last shape. Order the in-between shapes accordingly.
                    Shape prevShape = targetShape;
                    foreach (var shape in newlyAddedShapes)
                    {
                        Graphics.MoveZUntilBehind(shape, prevShape);
                        prevShape = shape;
                    }

                    if (newlyAddedShapes.Count >= dialog.YCopies)
                    {
                        for (int i = 0; i < dialog.YCopies; ++i)
                        {
                            Graphics.MoveZToJustInFront(newlyAddedShapes[i], targetShape);
                        }
                    }
                }
            }
            else
            {
                // Multiclone Between
                if (sourceShape.ZOrderPosition < targetShape.ZOrderPosition)
                {
                    // last shape in front of first shape. Order the in-between shapes accordingly.
                    foreach (var shape in newlyAddedShapes)
                    {
                        Graphics.MoveZUntilBehind(shape, targetShape);
                    }
                }
                else
                {
                    // first shape in front of last shape. Order the in-between shapes accordingly.
                    newlyAddedShapes.Reverse();
                    foreach (var shape in newlyAddedShapes)
                    {
                        Graphics.MoveZUntilBehind(shape, sourceShape);
                    }
                }
            }
        }


        public static void PivotAroundTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;
            var shapeList = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapeList.Count != 2)
            {
                Error(TextCollection.DrawingsLabSelectExactlyTwoShapes);
                return;
            }

            var sourceShape = shapeList[0];
            var pivotShape = shapeList[1];

            var dialog = new PivotAroundToolDialog(sourceShape, pivotShape);
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            double dx = dialog.SourceCenterX - dialog.PivotCenterX;
            double dy = dialog.SourceCenterY - dialog.PivotCenterY;
            double radius = Math.Sqrt(dx*dx + dy*dy);
            double initialAngle = Math.Atan2(dy, dx)*180/Math.PI;

            if (!dialog.FixOriginalLocation)
            {
                double radAngle = dialog.StartAngle*Math.PI/180;
                float cx = (float) (Math.Cos(radAngle)*radius + dialog.PivotCenterX);
                float cy = (float) (Math.Sin(radAngle)*radius + dialog.PivotCenterY);
                float anchorX = (float) dialog.SourceAnchorFractionX;
                float anchorY = (float) dialog.SourceAnchorFractionY;
                float angleDifference = (float) (dialog.StartAngle-initialAngle);
                
                Graphics.SetShapeX(sourceShape, cx, anchorX);
                Graphics.SetShapeY(sourceShape, cy, anchorY);
                if (dialog.RotateShape) Graphics.RotateShapeAboutPivot(sourceShape, angleDifference, anchorX, anchorY);
            }

            double angleStep = dialog.AngleDifference;
            if (!dialog.IsExtend) angleStep /= (dialog.Copies - 1);

            for (int i = 1; i < dialog.Copies; ++i)
            {
                var newShape = sourceShape.Duplicate()[1];
                double angle = dialog.StartAngle + angleStep*i;

                double radAngle = angle * Math.PI / 180;
                float cx = (float)(Math.Cos(radAngle) * radius + dialog.PivotCenterX);
                float cy = (float)(Math.Sin(radAngle) * radius + dialog.PivotCenterY);
                float anchorX = (float)dialog.SourceAnchorFractionX;
                float anchorY = (float)dialog.SourceAnchorFractionY;
                float angleDifference = (float) (angleStep*i);

                Graphics.SetShapeX(newShape, cx, anchorX);
                Graphics.SetShapeY(newShape, cy, anchorY);
                if (dialog.RotateShape) Graphics.RotateShapeAboutPivot(newShape, angleDifference, anchorX, anchorY);
            }
        }


        public static void SendBackward()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoSendBackward);
        }

        public static void BringForward()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoBringForward);
        }

        public static void SendToBack()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        public static void BringToFront()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoBringToFront);
        }

        public static void SendBehindShape()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count < 2)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapeToMoveBehind = shapes.Last();
            shapes.RemoveAt(shapes.Count - 1);

            Graphics.SortByZOrder(shapes);
            shapes.Reverse();
            foreach (var shape in shapes)
            {
                Graphics.MoveZUntilBehind(shape, shapeToMoveBehind);
            }
        }

        public static void BringInFrontOfShape()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count < 2)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var shapeToMoveInFront = shapes.Last();
            shapes.RemoveAt(shapes.Count - 1);

            Graphics.SortByZOrder(shapes);
            foreach (var shape in shapes)
            {
                Graphics.MoveZUntilInFront(shape, shapeToMoveInFront);
            }
        }

        public static void RecordDisplacement()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            var shapes = selection.ShapeRange;
            if (shapes.Count != 2)
            {
                Error(TextCollection.DrawingsLabSelectStartAndEndShape);
                return;
            }

            var firstShape = shapes[1];
            var secondShape = shapes[2];

            DataSource.ShiftValueX = GetX(secondShape) - GetX(firstShape);
            DataSource.ShiftValueY = GetY(secondShape) - GetY(firstShape);
            DataSource.ShiftValueRotation = secondShape.Rotation - firstShape.Rotation;
        }

        public static void ApplyDisplacement(bool applyAllSettings = false)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in selection.ShapeRange.Cast<Shape>())
            {
                if (applyAllSettings || DataSource.ShiftIncludePositionX)
                {
                    SetX(shape, GetX(shape) + DataSource.ShiftValueX);
                }
                if (applyAllSettings || DataSource.ShiftIncludePositionY)
                {
                    SetY(shape, GetY(shape) + DataSource.ShiftValueY);
                }
                if (applyAllSettings || DataSource.ShiftIncludeRotation)
                {
                    shape.Rotation += DataSource.ShiftValueRotation;
                }
            }
        }

        public static void RecordPosition()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            var shapes = selection.ShapeRange;
            if (shapes.Count != 1)
            {
                Error(TextCollection.DrawingsLabSelectExactlyOneShape);
                return;
            }
            var shape = shapes[1];

            DataSource.SavedValueX = GetX(shape);
            DataSource.SavedValueY = GetY(shape);
            DataSource.SavedValueRotation = shape.Rotation;
        }

        public static void ApplyPosition(bool applyAllSettings = false)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in selection.ShapeRange.Cast<Shape>())
            {
                if (applyAllSettings || DataSource.SavedIncludePositionX)
                {
                    SetX(shape, DataSource.SavedValueX);
                }
                if (applyAllSettings || DataSource.SavedIncludePositionY)
                {
                    SetY(shape, DataSource.SavedValueY);
                }
                if (applyAllSettings || DataSource.SavedIncludeRotation)
                {
                    shape.Rotation = DataSource.SavedValueRotation;
                }
            }
        }


        public static void RecordFormat()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            var shapes = selection.ShapeRange;
            if (shapes.Count != 1)
            {
                Error(TextCollection.DrawingsLabSelectExactlyOneShape);
                return;
            }
            var shape = shapes[1];

            try
            {
                var font = shape.TextFrame2.TextRange.Font;
                DataSource.FormatText = Graphics.GetText(shape);
                DataSource.FormatTextColor = font.Fill.ForeColor.RGB;
                DataSource.FormatTextFontSize = font.Size;
                DataSource.FormatTextFont = font.Name;
                DataSource.FormatTextWrap = shape.TextFrame2.WordWrap == MsoTriState.msoTrue;
                DataSource.FormatTextAutoSize = shape.TextFrame2.AutoSize;
            }
            catch (ArgumentException e)
            {
                // ArgumentException is thrown if the shape does not have this property.
            }

            try
            {
                var line = shape.Line;
                DataSource.FormatHasLine = line.Visible == MsoTriState.msoTrue;
                DataSource.FormatLineColor = line.ForeColor.RGB;
                DataSource.FormatLineWeight = line.Weight;
                DataSource.FormatLineDashStyle = line.DashStyle;
            }
            catch (ArgumentException e)
            {
                // ArgumentException is thrown if the shape does not have this property.
            }

            try
            {
                var fill = shape.Fill;
                DataSource.FormatHasFill = fill.Visible == MsoTriState.msoTrue;
                DataSource.FormatFillColor = fill.ForeColor.RGB;
            }
            catch (ArgumentException e)
            {
                // ArgumentException is thrown if the shape does not have this property.
            }

            DataSource.FormatWidth = shape.Width;
            DataSource.FormatHeight = shape.Height;
        }

        public static void ApplyFormat(bool applyAllSettings = false)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();

            Action<bool, bool, Action> apply = (isDefaultSetting, condition, action) =>
            {
                if (applyAllSettings && !isDefaultSetting) return;
                if (!applyAllSettings && !condition) return;

                try
                {
                    action();
                }
                catch (ArgumentException e)
                {
                    // ArgumentException is thrown if the shape does not have this property.
                }
            };

            foreach (var s in selection.ShapeRange.Cast<Shape>())
            {
                var shape = s;

                // Sync Text Style
                apply(false, DataSource.FormatSyncTextStyle && DataSource.FormatIncludeText,
                    () => Graphics.SetText(shape, DataSource.FormatText));
                apply(true, DataSource.FormatSyncTextStyle && DataSource.FormatIncludeTextColor,
                    () => shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = DataSource.FormatTextColor);
                apply(true, DataSource.FormatSyncTextStyle && DataSource.FormatIncludeTextFontSize,
                    () => shape.TextFrame2.TextRange.Font.Size = DataSource.FormatTextFontSize);
                apply(true, DataSource.FormatSyncTextStyle && DataSource.FormatIncludeTextFont,
                    () => shape.TextFrame2.TextRange.Font.Name = DataSource.FormatTextFont);
                apply(true, DataSource.FormatSyncTextStyle && DataSource.FormatIncludeTextWrap,
                    () => shape.TextFrame2.WordWrap = DataSource.FormatTextWrap ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                apply(true, DataSource.FormatSyncTextStyle && DataSource.FormatIncludeTextAutoSize,
                    () => shape.TextFrame2.AutoSize = DataSource.FormatTextAutoSize);

                // Sync Line Style
                apply(true, DataSource.FormatSyncLineStyle && DataSource.FormatIncludeHasLine,
                    () => shape.Line.Visible = DataSource.FormatHasLine ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                apply(true, DataSource.FormatSyncLineStyle && DataSource.FormatIncludeLineColor,
                    () => shape.Line.ForeColor.RGB = DataSource.FormatLineColor);
                apply(true, DataSource.FormatSyncLineStyle && DataSource.FormatIncludeLineWeight,
                    () => shape.Line.Weight = DataSource.FormatLineWeight);
                apply(true, DataSource.FormatSyncLineStyle && DataSource.FormatIncludeLineDashStyle,
                    () => shape.Line.DashStyle = DataSource.FormatLineDashStyle);

                // Sync Fill Style
                apply(true, DataSource.FormatSyncFillStyle && DataSource.FormatIncludeHasFill,
                    () => shape.Fill.Visible = DataSource.FormatHasFill ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                apply(true, DataSource.FormatSyncFillStyle && DataSource.FormatIncludeFillColor,
                    () => shape.Fill.ForeColor.RGB = DataSource.FormatFillColor);

                // Sync Size
                apply(false, DataSource.FormatSyncSize && DataSource.FormatIncludeWidth,
                    () => shape.Width = DataSource.FormatWidth);
                apply(false, DataSource.FormatSyncSize && DataSource.FormatIncludeHeight,
                    () => shape.Height = DataSource.FormatHeight);
            }
        }

        public static void SetControlGroup(Native.VirtualKey key, bool appendToGroup = false)
        {
            if (!Native.IsNumberKey(key)) return;
            if (appendToGroup)
            {
                SelectControlGroup(key, true);
            }

            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            var shapes = new HashSet<int>(selection.ShapeRange.Cast<Shape>().Select(shape => shape.Id));
            var slideId = PowerPointCurrentPresentationInfo.CurrentSlide.ID;

            ControlGroups[key] = new ControlGroup(slideId, shapes);
        }

        public static void SelectControlGroup(Native.VirtualKey key, bool appendToSelection = false)
        {
            if (!Native.IsNumberKey(key)) return;

            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionSlides) return;

            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (!ControlGroups.ContainsKey(key)) return;

            var controlGroup = ControlGroups[key];
            var targetSlide = PowerPointPresentation.Current.Slides.FirstOrDefault(slide => slide.ID == controlGroup.SlideId);
            if (targetSlide == null) return;


            targetSlide.GetNativeSlide().Select();

            if (!appendToSelection)
            {
                Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            }

            var shapeIds = controlGroup.ShapeIds;
            var shapes = currentSlide.Shapes.Cast<Shape>()
                                            .Where(shape => shapeIds.Contains(shape.Id));
            foreach (var shape in shapes)
            {
                shape.Visible = MsoTriState.msoTrue;
                shape.Select(MsoTriState.msoFalse);
            }
        }


        public static void SelectAllOfType()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes) return;

            var selectedShapeTypes = new HashSet<MsoAutoShapeType>(selection.ShapeRange.Cast<Shape>().Select(shape => shape.AutoShapeType));

            PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Cast<Shape>()
                                                                 .Where(shape => selectedShapeTypes.Contains(shape.AutoShapeType))
                                                                 .ToList()
                                                                 .ForEach(shape => shape.Select(MsoTriState.msoFalse));
        }

        public static void AlignHorizontal()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }
            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count <= 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var dialog = new AlignmentDialogHorizontal();
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchor = dialog.SourceAnchor / 100;
            double targetAnchor = dialog.TargetAnchor / 100;

            var targetShape = shapes.Last();
            shapes.RemoveAt(shapes.Count - 1);

            double targetY = targetShape.Top + (1 - targetAnchor) * targetShape.Height;
            foreach (var shape in shapes)
            {
                shape.Top = (float)(targetY - (1 - sourceAnchor) * shape.Height);
            }
        }

        public static void AlignVertical()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }
            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count <= 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var dialog = new AlignmentDialogVertical();
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchor = dialog.SourceAnchor / 100;
            double targetAnchor = dialog.TargetAnchor / 100;

            var targetShape = shapes.Last();
            shapes.RemoveAt(shapes.Count - 1);

            double targetX = targetShape.Left + targetAnchor * targetShape.Width;
            foreach (var shape in shapes)
            {
                shape.Left = (float)(targetX - sourceAnchor * shape.Width);
            }
        }

        public static void AlignBoth()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }
            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count <= 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var dialog = new AlignmentDialogBoth();
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchorX = dialog.SourceAnchorVertical / 100;
            double targetAnchorX = dialog.TargetAnchorVertical / 100;
            double sourceAnchorY = dialog.SourceAnchorHorizontal / 100;
            double targetAnchorY = dialog.TargetAnchorHorizontal / 100;

            var targetShape = shapes.Last();
            shapes.RemoveAt(shapes.Count - 1);

            double targetX = targetShape.Left + targetAnchorX * targetShape.Width;
            double targetY = targetShape.Top + (1 - targetAnchorY) * targetShape.Height;
            foreach (var shape in shapes)
            {
                shape.Left = (float)(targetX - sourceAnchorX * shape.Width);
                shape.Top = (float)(targetY - (1 - sourceAnchorY) * shape.Height);
            }
        }


        public static void AlignVerticalToSlide()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }
            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count < 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            var dialog = new AlignmentDialogHorizontal();
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchor = dialog.SourceAnchor / 100;
            double targetAnchor = dialog.TargetAnchor / 100;

            double targetY = (1 - targetAnchor) * PowerPointPresentation.Current.SlideHeight;
            foreach (var shape in shapes)
            {
                shape.Top = (float)(targetY - (1 - sourceAnchor) * shape.Height);
            }
        }

        public static void AlignHorizontalToSlide()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }
            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count < 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            var dialog = new AlignmentDialogVertical();
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchor = dialog.SourceAnchor / 100;
            double targetAnchor = dialog.TargetAnchor / 100;

            double targetX = targetAnchor * PowerPointPresentation.Current.SlideWidth;
            foreach (var shape in shapes)
            {
                shape.Left = (float)(targetX - sourceAnchor * shape.Width);
            }
        }


        public static void AlignBothToSlide()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }
            var shapes = selection.ShapeRange.Cast<Shape>().ToList();
            if (shapes.Count < 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            var dialog = new AlignmentDialogBoth();
            if (dialog.ShowDialog() != true) return;
            if (dialog.DialogResult != true) return;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchorX = dialog.SourceAnchorVertical / 100;
            double targetAnchorX = dialog.TargetAnchorVertical / 100;
            double sourceAnchorY = dialog.SourceAnchorHorizontal / 100;
            double targetAnchorY = dialog.TargetAnchorHorizontal / 100;

            double targetX = targetAnchorX * PowerPointPresentation.Current.SlideWidth;
            double targetY = (1 - targetAnchorY) * PowerPointPresentation.Current.SlideHeight;
            foreach (var shape in shapes)
            {
                shape.Left = (float)(targetX - sourceAnchorX * shape.Width);
                shape.Top = (float)(targetY - (1 - sourceAnchorY) * shape.Height);
            }
        }

        #endregion

        #region Convenience Functions
        public static float GetX(Shape shape)
        {
            switch (DataSource.AnchorHorizontal)
            {
                case DrawingsLabDataSource.Horizontal.Left:
                    return shape.Left;
                case DrawingsLabDataSource.Horizontal.Center:
                    return Graphics.GetMidpointX(shape);
                case DrawingsLabDataSource.Horizontal.Right:
                    return Graphics.GetRight(shape);
            }
            throw new ArgumentOutOfRangeException();
        }

        public static void SetX(Shape shape, float value)
        {
            switch (DataSource.AnchorHorizontal)
            {
                case DrawingsLabDataSource.Horizontal.Left:
                    shape.Left = value;
                    return;
                case DrawingsLabDataSource.Horizontal.Center:
                    Graphics.SetMidpointX(shape, value);
                    return;
                case DrawingsLabDataSource.Horizontal.Right:
                    Graphics.SetRight(shape, value);
                    return;
            }
            throw new ArgumentOutOfRangeException();
        }

        public static float GetY(Shape shape)
        {
            switch (DataSource.AnchorVertical)
            {
                case DrawingsLabDataSource.Vertical.Top:
                    return shape.Top;
                case DrawingsLabDataSource.Vertical.Middle:
                    return Graphics.GetMidpointY(shape);
                case DrawingsLabDataSource.Vertical.Bottom:
                    return Graphics.GetBottom(shape);
            }
            throw new ArgumentOutOfRangeException();
        }

        public static void SetY(Shape shape, float value)
        {
            switch (DataSource.AnchorVertical)
            {
                case DrawingsLabDataSource.Vertical.Top:
                    shape.Top = value;
                    return;
                case DrawingsLabDataSource.Vertical.Middle:
                    Graphics.SetMidpointY(shape, value);
                    return;
                case DrawingsLabDataSource.Vertical.Bottom:
                    Graphics.SetBottom(shape, value);
                    return;
            }
            throw new ArgumentOutOfRangeException();
        }

        #endregion


        #region Utility Functions
        private static void Error(string message)
        {
            MessageBox.Show(message, "Error");
            // for now do nothing.
        }

        #endregion
    }
}
