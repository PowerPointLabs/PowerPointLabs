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
using PowerPointLabs.Views;
using PPExtraEventHelper;

using Graphics = PowerPointLabs.Utils.Graphics;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.DrawingsLab
{
    internal class DrawingsLabMain
    {
#pragma warning disable 0618
        private readonly DrawingsLabDataSource _dataSource;
        private readonly Dictionary<Native.VirtualKey, ControlGroup> _controlGroups = new Dictionary<Native.VirtualKey, ControlGroup>();  

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

        public DrawingsLabMain(DrawingLabData data)
        {
            _dataSource = new DrawingsLabDataSource();
            _dataSource.AssignData(data);
        }

        public Action FunctionWrapper(Action action)
        {
            return () =>
            {
                try
                {
                    action();
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog("Unexpected error in drawings lab.", e.Message, e);
                    throw e;
                }
                GC.Collect();
            };
        }


        #region API

        public void SwitchToLineTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeStraightConnector");
        }

        public void SwitchToArrowTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeStraightConnectorArrow");
        }

        public void SwitchToRectangleTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeRectangle");
        }

        public void SwitchToRoundedRectangleTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeRoundedRectangle");
        }

        public void SwitchToCircleTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeOval");
        }

        public void SwitchToTriangleTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("ShapeIsoscelesTriangle");
        }

        public void SwitchToTextboxTool()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;

            // Either one or the other will work. If TextBoxInsert doesn't work, use TextBoxInsertHorizontal instead.
            try
            {
                commandBars.ExecuteMso("TextBoxInsert");
            }
            catch (COMException)
            {
                commandBars.ExecuteMso("TextBoxInsertHorizontal");
            }
        }

        public void AddText()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            var text = DrawingsLabDialogs.ShowInsertTextDialog();
            if (text == null)
            {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes)
            {
                try
                {
                    Graphics.SetText(shape, text);
                }
                catch (ArgumentException)
                {
                    Debug.WriteLine("Unable to write text to " + shape.Name);
                }
            }
        }

        public void AddMath()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count != 1)
            {
                Error(TextCollection.DrawingsLabSelectExactlyOneShape);
                return;
            }

            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                var commandBars = Globals.ThisAddIn.Application.CommandBars;
                commandBars.ExecuteMso("EquationInsertNew");
            }
            catch (COMException)
            {
                // Do nothing. EquationInsertNew throws an exception even as it succeeds.
            }
        }

        public void RemoveText()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes)
            {
                Graphics.SetText(shape, String.Empty);
            }
        }

        public void GroupShapes()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count < 2)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var slide = PowerPointCurrentPresentationInfo.CurrentSlide;

            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                slide.GroupShapes(shapes);
            }
            catch (UnauthorizedAccessException)
            {
                Error(TextCollection.DrawingsLabErrorCannotGroup);
            }
        }

        public void UngroupShapes()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }
            
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            bool didSomething = false;
            foreach (var shape in shapes.Where(Graphics.IsAGroup))
            {
                shape.Ungroup();
                didSomething = true;
            }
            if (!didSomething)
            {
                Error(TextCollection.DrawingsLabErrorNothingUngrouped);
            }
        }

        public void ToggleArrowStart()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var allArrowHeads = shapes.Where(Graphics.CanAddArrows)
                                      .All(shape => shape.Line.BeginArrowheadStyle != MsoArrowheadStyle.msoArrowheadNone);

            if (allArrowHeads)
            {
                foreach (var shape in shapes.Where(Graphics.CanAddArrows))
                {
                    shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
                }
            }
            else
            {
                foreach (var shape in shapes.Where(Graphics.CanAddArrows)
                    .Where(shape => shape.Line.BeginArrowheadStyle == MsoArrowheadStyle.msoArrowheadNone))
                {
                    shape.Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadOpen;
                }
            }
        }

        public void ToggleArrowEnd()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            var allArrowHeads = shapes.Where(Graphics.CanAddArrows)
                                      .All(shape => shape.Line.EndArrowheadStyle != MsoArrowheadStyle.msoArrowheadNone);

            if (allArrowHeads)
            {
                foreach (var shape in shapes.Where(Graphics.CanAddArrows))
                {
                    shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone;
                }
            }
            else
            {
                foreach (var shape in shapes.Where(Graphics.CanAddArrows)
                    .Where(shape => shape.Line.EndArrowheadStyle == MsoArrowheadStyle.msoArrowheadNone))
                {
                    shape.Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadOpen;
                }
            }
        }


        public void HideTool()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes)
            {
                shape.Visible = MsoTriState.msoFalse;
            }
        }

        public void ShowAllTool()
        {
            var shapes = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes.Cast<Shape>())
            {
                shape.Visible = MsoTriState.msoTrue;
            }
        }

        public void OpenSelectionPane()
        {
            var commandBars = Globals.ThisAddIn.Application.CommandBars;
            commandBars.ExecuteMso("SelectionPane");
        }

        public void CloneTool()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            PowerPointCurrentPresentationInfo.CurrentSlide.CopyShapesToSlide(selection.ShapeRange);
        }

        public void MultiCloneExtendTool()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count == 0 || shapes.Count % 2 != 0)
            {
                Error(TextCollection.DrawingsLabSelectTwoSetsOfShapes);
                return;
            }

            int clones = DrawingsLabDialogs.ShowMultiCloneNumericDialog();
            if (clones <= 0)
            {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            int midpoint = shapes.Count / 2;
            for (int i = 0; i < shapes.Count / 2; ++i)
            {
                // Do the cloning for every pair of shapes (i, midpoint+i)
                var firstShape = shapes[i];
                var secondShape = shapes[midpoint + i];

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

        public void MultiCloneBetweenTool()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count == 0 || shapes.Count % 2 != 0)
            {
                Error(TextCollection.DrawingsLabSelectTwoSetsOfShapes);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            int clones = DrawingsLabDialogs.ShowMultiCloneNumericDialog();
            if (clones <= 0)
            {
                return;
            }

            int divisions = clones + 1;

            int midpoint = shapes.Count / 2;
            for (int i = 0; i < shapes.Count / 2; ++i)
            {
                // Do the cloning for every pair of shapes (i, midpoint+i)
                var firstShape = shapes[i];
                var lastShape = shapes[midpoint + i];

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

        public void MultiCloneGridTool()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count != 2)
            {
                Error(TextCollection.DrawingsLabSelectExactlyTwoShapes);
                return;
            }

            var sourceShape = shapes[0];
            var targetShape = shapes[1];

            var dialog = new MultiCloneGridDialog(sourceShape.Left, sourceShape.Top, targetShape.Left, targetShape.Top);
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

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
                    if (x == 0 && y == 0)
                    {
                        continue;
                    }

                    if (x == skipIndexX && y == skipIndexY)
                    {
                        continue;
                    }

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


        public void PivotAroundTool()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count != 2)
            {
                Error(TextCollection.DrawingsLabSelectExactlyTwoShapes);
                return;
            }

            var sourceShape = shapes[0];
            var pivotShape = shapes[1];

            var dialog = new PivotAroundToolDialog(sourceShape, pivotShape);
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

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
                if (dialog.RotateShape)
                {
                    Graphics.RotateShapeAboutPivot(sourceShape, angleDifference, anchorX, anchorY);
                }
            }

            double angleStep = dialog.AngleDifference;
            if (!dialog.IsExtend)
            {
                angleStep /= (dialog.Copies - 1);
            }

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
                if (dialog.RotateShape)
                {
                    Graphics.RotateShapeAboutPivot(newShape, angleDifference, anchorX, anchorY);
                }
            }
        }


        public void SendBackward()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoSendBackward);
        }

        public void BringForward()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoBringForward);
        }

        public void SendToBack()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoSendToBack);
        }

        public void BringToFront()
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            if (selection.Type != PpSelectionType.ppSelectionShapes && selection.Type != PpSelectionType.ppSelectionText)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            selection.ShapeRange.ZOrder(MsoZOrderCmd.msoBringToFront);
        }

        public void SendBehindShape()
        {
            var shapes = GetCurrentlySelectedShapes();
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

        public void BringInFrontOfShape()
        {
            var shapes = GetCurrentlySelectedShapes();
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

        public void RecordDisplacement()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count != 2)
            {
                Error(TextCollection.DrawingsLabSelectStartAndEndShape);
                return;
            }

            var firstShape = shapes[0];
            var secondShape = shapes[1];

            _dataSource.ShiftValueX = GetX(secondShape) - GetX(firstShape);
            _dataSource.ShiftValueY = GetY(secondShape) - GetY(firstShape);
            _dataSource.ShiftValueRotation = secondShape.Rotation - firstShape.Rotation;
        }

        public void ApplyDisplacement(bool applyAllSettings = false)
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes)
            {
                if (applyAllSettings || _dataSource.ShiftIncludePositionX)
                {
                    SetX(shape, GetX(shape) + _dataSource.ShiftValueX);
                }
                if (applyAllSettings || _dataSource.ShiftIncludePositionY)
                {
                    SetY(shape, GetY(shape) + _dataSource.ShiftValueY);
                }
                if (applyAllSettings || _dataSource.ShiftIncludeRotation)
                {
                    shape.Rotation += _dataSource.ShiftValueRotation;
                }
            }
        }

        public void RecordPosition()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count != 1)
            {
                Error(TextCollection.DrawingsLabSelectExactlyOneShape);
                return;
            }
            var shape = shapes[0];

            _dataSource.SavedValueX = GetX(shape);
            _dataSource.SavedValueY = GetY(shape);
            _dataSource.SavedValueRotation = shape.Rotation;
        }

        public void ApplyPosition(bool applyAllSettings = false)
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            foreach (var shape in shapes)
            {
                if (applyAllSettings || _dataSource.SavedIncludePositionX)
                {
                    SetX(shape, _dataSource.SavedValueX);
                }
                if (applyAllSettings || _dataSource.SavedIncludePositionY)
                {
                    SetY(shape, _dataSource.SavedValueY);
                }
                if (applyAllSettings || _dataSource.SavedIncludeRotation)
                {
                    shape.Rotation = _dataSource.SavedValueRotation;
                }
            }
        }


        public void RecordFormat()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count != 1)
            {
                Error(TextCollection.DrawingsLabSelectExactlyOneShape);
                return;
            }
            var shape = shapes[0];

            try
            {
                var font = shape.TextFrame2.TextRange.Font;
                _dataSource.FormatText = Graphics.GetText(shape);
                _dataSource.FormatTextColor = font.Fill.ForeColor.RGB;
                _dataSource.FormatTextFontSize = font.Size;
                _dataSource.FormatTextFont = font.Name;
                _dataSource.FormatTextWrap = shape.TextFrame2.WordWrap == MsoTriState.msoTrue;
                _dataSource.FormatTextAutoSize = shape.TextFrame2.AutoSize;
            }
            catch (ArgumentException)
            {
                // ArgumentException is thrown if the shape does not have this property.
            }

            try
            {
                var line = shape.Line;
                _dataSource.FormatHasLine = line.Visible == MsoTriState.msoTrue;
                _dataSource.FormatLineColor = line.ForeColor.RGB;
                _dataSource.FormatLineWeight = line.Weight;
                _dataSource.FormatLineDashStyle = line.DashStyle;
            }
            catch (ArgumentException)
            {
                // ArgumentException is thrown if the shape does not have this property.
            }

            try
            {
                var fill = shape.Fill;
                _dataSource.FormatHasFill = fill.Visible == MsoTriState.msoTrue;
                _dataSource.FormatFillColor = fill.ForeColor.RGB;
            }
            catch (ArgumentException)
            {
                // ArgumentException is thrown if the shape does not have this property.
            }

            _dataSource.FormatWidth = shape.Width;
            _dataSource.FormatHeight = shape.Height;
        }

        public void ApplyFormat(bool applyAllSettings = false)
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();

            Action<bool, bool, Action> apply = (isDefaultSetting, condition, action) =>
            {
                if (applyAllSettings && !isDefaultSetting)
                {
                    return;
                }

                if (!applyAllSettings && !condition)
                {
                    return;
                }

                try
                {
                    action();
                }
                catch (ArgumentException)
                {
                    // ArgumentException is thrown if the shape does not have this property.
                }
            };

            foreach (var s in shapes)
            {
                var shape = s;

                // Sync Text Style
                apply(false, _dataSource.FormatSyncTextStyle && _dataSource.FormatIncludeText,
                    () => Graphics.SetText(shape, _dataSource.FormatText));
                apply(true, _dataSource.FormatSyncTextStyle && _dataSource.FormatIncludeTextColor,
                    () => shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _dataSource.FormatTextColor);
                apply(true, _dataSource.FormatSyncTextStyle && _dataSource.FormatIncludeTextFontSize,
                    () => shape.TextFrame2.TextRange.Font.Size = _dataSource.FormatTextFontSize);
                apply(true, _dataSource.FormatSyncTextStyle && _dataSource.FormatIncludeTextFont,
                    () => shape.TextFrame2.TextRange.Font.Name = _dataSource.FormatTextFont);
                apply(true, _dataSource.FormatSyncTextStyle && _dataSource.FormatIncludeTextWrap,
                    () => shape.TextFrame2.WordWrap = _dataSource.FormatTextWrap ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                apply(true, _dataSource.FormatSyncTextStyle && _dataSource.FormatIncludeTextAutoSize,
                    () => shape.TextFrame2.AutoSize = _dataSource.FormatTextAutoSize);

                // Sync Line Style
                apply(true, _dataSource.FormatSyncLineStyle && _dataSource.FormatIncludeHasLine,
                    () => shape.Line.Visible = _dataSource.FormatHasLine ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                apply(true, _dataSource.FormatSyncLineStyle && _dataSource.FormatIncludeLineColor,
                    () => shape.Line.ForeColor.RGB = _dataSource.FormatLineColor);
                apply(true, _dataSource.FormatSyncLineStyle && _dataSource.FormatIncludeLineWeight,
                    () => shape.Line.Weight = _dataSource.FormatLineWeight);
                apply(true, _dataSource.FormatSyncLineStyle && _dataSource.FormatIncludeLineDashStyle,
                    () => shape.Line.DashStyle = _dataSource.FormatLineDashStyle);

                // Sync Fill Style
                apply(true, _dataSource.FormatSyncFillStyle && _dataSource.FormatIncludeHasFill,
                    () => shape.Fill.Visible = _dataSource.FormatHasFill ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                apply(true, _dataSource.FormatSyncFillStyle && _dataSource.FormatIncludeFillColor,
                    () => shape.Fill.ForeColor.RGB = _dataSource.FormatFillColor);

                // Sync Size
                apply(false, _dataSource.FormatSyncSize && _dataSource.FormatIncludeWidth,
                    () => shape.Width = _dataSource.FormatWidth);
                apply(false, _dataSource.FormatSyncSize && _dataSource.FormatIncludeHeight,
                    () => shape.Height = _dataSource.FormatHeight);
            }
        }

        public void SetControlGroup(Native.VirtualKey key, bool appendToGroup = false)
        {
            if (!Native.IsNumberKey(key))
            {
                return;
            }

            if (appendToGroup)
            {
                SelectControlGroup(key, true);
            }

            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            var shapes = new HashSet<int>(selection.ShapeRange.Cast<Shape>().Select(shape => shape.Id));
            var slideId = PowerPointCurrentPresentationInfo.CurrentSlide.ID;

            _controlGroups[key] = new ControlGroup(slideId, shapes);
        }

        public void SelectControlGroup(Native.VirtualKey key, bool appendToSelection = false)
        {
            if (!Native.IsNumberKey(key))
            {
                return;
            }

            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionSlides)
            {
                return;
            }

            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (!_controlGroups.ContainsKey(key))
            {
                return;
            }

            var controlGroup = _controlGroups[key];
            var targetSlide = PowerPointPresentation.Current.Slides.FirstOrDefault(slide => slide.ID == controlGroup.SlideId);
            if (targetSlide == null)
            {
                return;
            }


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


        public void SelectAllOfType()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            var selectedShapeTypes = new HashSet<MsoAutoShapeType>(selection.ShapeRange.Cast<Shape>().Select(shape => shape.AutoShapeType));

            PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Cast<Shape>()
                                                                 .Where(shape => selectedShapeTypes.Contains(shape.AutoShapeType))
                                                                 .ToList()
                                                                 .ForEach(shape => shape.Select(MsoTriState.msoFalse));
        }

        public void AlignHorizontal()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var dialog = new AlignmentDialogHorizontal();
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

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

        public void AlignVertical()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var dialog = new AlignmentDialogVertical();
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

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

        public void AlignBoth()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 1)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastTwoShapes);
                return;
            }

            var dialog = new AlignmentDialogBoth();
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

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


        public void AlignVerticalToSlide()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            var dialog = new AlignmentDialogHorizontal();
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchor = dialog.SourceAnchor / 100;
            double targetAnchor = dialog.TargetAnchor / 100;

            double targetY = (1 - targetAnchor) * PowerPointPresentation.Current.SlideHeight;
            foreach (var shape in shapes)
            {
                shape.Top = (float)(targetY - (1 - sourceAnchor) * shape.Height);
            }
        }

        public void AlignHorizontalToSlide()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            var dialog = new AlignmentDialogVertical();
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

            Globals.ThisAddIn.Application.StartNewUndoEntry();
            double sourceAnchor = dialog.SourceAnchor / 100;
            double targetAnchor = dialog.TargetAnchor / 100;

            double targetX = targetAnchor * PowerPointPresentation.Current.SlideWidth;
            foreach (var shape in shapes)
            {
                shape.Left = (float)(targetX - sourceAnchor * shape.Width);
            }
        }


        public void AlignBothToSlide()
        {
            var shapes = GetCurrentlySelectedShapes();
            if (shapes.Count <= 0)
            {
                Error(TextCollection.DrawingsLabSelectAtLeastOneShape);
                return;
            }

            var dialog = new AlignmentDialogBoth();
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            if (dialog.DialogResult != true)
            {
                return;
            }

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
        private float GetX(Shape shape)
        {
            switch (_dataSource.AnchorHorizontal)
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

        private void SetX(Shape shape, float value)
        {
            switch (_dataSource.AnchorHorizontal)
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

        private float GetY(Shape shape)
        {
            switch (_dataSource.AnchorVertical)
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

        private void SetY(Shape shape, float value)
        {
            switch (_dataSource.AnchorVertical)
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

        private List<Shape> GetCurrentlySelectedShapes()
        {
            var selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
            {
                return selection.ShapeRange.Cast<Shape>().ToList();
            }
            return new List<Shape>();
        }

        #endregion


        #region Utility Functions
        private void Error(string message)
        {
            MessageBox.Show(message, "Error");
            // for now do nothing.
        }

        #endregion
    }
}
