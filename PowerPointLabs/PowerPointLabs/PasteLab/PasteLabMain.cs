﻿using System;
using System.Collections.Generic;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PasteLab
{
    public class PasteLabMain
    {
        public static void PasteToFillSlide(PowerPointSlide slide, bool clipboardIsEmpty, float width, float height)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToFillSlide encountered empty clipboard");
                return;
            }

            ShapeRange pastedShapeRange = slide.Shapes.Paste();
            Logger.Log(string.Format("PasteToFillSlide: {0} objects pasted", pastedShapeRange.Count));
            pastedShapeRange = RemovePlaceholders(slide, pastedShapeRange);

            if (pastedShapeRange.Count <= 0)
            {
                Logger.Log("No resizable objects, PasteToFillSlide finished early");
                return;
            }

            var resizeShape = pastedShapeRange[1];
            if (pastedShapeRange.Count > 1)
            {
                resizeShape = pastedShapeRange.Group();
            }
            resizeShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            var ppResizeShape = new PPShape(resizeShape);
            
            ppResizeShape.AbsoluteHeight = height;
            if (ppResizeShape.AbsoluteWidth < width)
            {
                ppResizeShape.AbsoluteWidth = width;
            }
            ppResizeShape.VisualCenter = new System.Drawing.PointF(width / 2, height / 2);
            
            CropLab.CropToSlide.Crop(resizeShape, slide, width, height);
        }

        public static void PasteAndReplace(PowerPointPresentation presentation, PowerPointSlide slide,
                                           bool clipboardIsEmpty, Selection selection)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteAndReplace encountered an empty clipboard");
                return;
            }

            if (selection.ShapeRange.Count == 0)
            {
                Logger.Log("PasteAndReplace found no shapes selected");
                return;
            }

            var shapeToReplace = selection.ShapeRange[1];

            if (selection.HasChildShapeRange)
            {
                Logger.Log("PasteAndReplace: Replacing item in group");
                shapeToReplace = selection.ChildShapeRange[1];
                selection.ShapeRange[1].Select();
                var pastedShapes = PasteIntoGroup(presentation, slide, clipboardIsEmpty, selection);
                pastedShapes.Left = shapeToReplace.Left;
                pastedShapes.Top = shapeToReplace.Top;
                shapeToReplace.Delete();
                return;
            }

            Shape newShape = slide.Shapes.Paste()[1];
            newShape.Left = shapeToReplace.Left;
            newShape.Top = shapeToReplace.Top;

            foreach (Effect eff in slide.TimeLine.MainSequence)
            {
                if (eff.Shape == shapeToReplace)
                {
                    Effect newEff = slide.TimeLine.MainSequence.Clone(eff);
                    newEff.Shape = newShape;
                    eff.Delete();
                }
            }

            shapeToReplace.PickUp();
            newShape.Apply();

            Logger.Log(string.Format("PasteAndReplace: Replaced {0} with {1}", shapeToReplace.Name, newShape.Name));
            shapeToReplace.Delete();
        }

        public static ShapeRange PasteIntoGroup(PowerPointPresentation presentation, PowerPointSlide slide,
                                                bool clipboardIsEmpty, Selection selection)
        {
            var newSlide = presentation.AddSlide();
            var selectedShapes = selection.ShapeRange;

            ShapeRange pastedShapes = slide.Shapes.Paste();

            selection.Copy();
            newSlide.Shapes.Paste();
            pastedShapes.Copy();

            List<int> order = new List<int>();

            foreach (Effect eff in slide.TimeLine.MainSequence)
            {
                if (eff.Shape.Equals(selectedShapes[1]))
                {
                    order.Add(eff.Index);
                }
            }

            selectedShapes = selectedShapes.Ungroup();


            List<String> newShapeNames = new List<String>();

            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                newShapeNames.Add(shape.Name);
            }

            foreach (PowerPoint.Shape shape in pastedShapes)
            {
                newShapeNames.Add(shape.Name);
            }

            PowerPoint.ShapeRange newShapeRange = slide.Shapes.Range(newShapeNames.ToArray());
            PowerPoint.Shape newGroupedShape = newShapeRange.Group();

            TransferEffects(order, newGroupedShape, slide, newSlide);

            newSlide.Delete();

            return pastedShapes;
        }

        public static void GroupSelectedShapes(PowerPointPresentation presentation, Models.PowerPointSlide slide,
                                               PowerPoint.Selection selection)
        {
            if (selection.ShapeRange.Count < 2)
            {
                MessageBox.Show("Please select more than one shape.", "Error");
                return;
            }

            var newSlide = presentation.AddSlide();
            var selectedShapes = selection.ShapeRange;

            selectedShapes[1].Copy();
            newSlide.Shapes.Paste();

            List<int> order = new List<int>();

            foreach (Effect eff in slide.TimeLine.MainSequence)
            {
                if (eff.Shape.Equals(selectedShapes[1]))
                {
                    order.Add(eff.Index);
                }
            }

            PowerPoint.Shape newGroupedShape = selectedShapes.Group();

            TransferEffects(order, newGroupedShape, slide, newSlide);

            newSlide.Delete();
        }

        public static void PasteToPosition(Models.PowerPointSlide slide, bool clipboardIsEmpty, float xPosition, float yPosition)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToPosition encountered an empty clipboard");
                return;
            }

            var pastedShapeRange = slide.Shapes.Paste();
            pastedShapeRange = RemovePlaceholders(slide, pastedShapeRange);

            if (pastedShapeRange.Count > 1)
            {
                Shape pastedShapeGroup = pastedShapeRange.Group();
                pastedShapeGroup.Left = xPosition;
                pastedShapeGroup.Top = yPosition;
                Logger.Log(string.Format("PasteToPosition: Pasted {0} at ({1}, {2})", pastedShapeGroup.Name, pastedShapeGroup.Left, pastedShapeGroup.Top));
                pastedShapeGroup.Ungroup();
            }
            else if (pastedShapeRange.Count == 1)
            {
                pastedShapeRange.Left = xPosition;
                pastedShapeRange.Top = yPosition;
                Logger.Log(string.Format("PasteToPosition: Pasted {0} at ({1}, {2})", pastedShapeRange.Name, pastedShapeRange.Left, pastedShapeRange.Top));
            }
        }

        public static void PasteToOriginalPosition(PowerPointPresentation presentation,
                                                   PowerPointSlide slide, bool clipboardIsEmpty)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToOriginalPosition encountered an empty clipboard");
                return;
            }

            var newSlide = presentation.AddSlide();

            ShapeRange correctShapes = newSlide.Shapes.Paste();

            foreach (Shape shape in correctShapes)
            {
                shape.Copy();
                Shape pastedShape = slide.Shapes.Paste()[1];
                pastedShape.Top = shape.Top;
                pastedShape.Left = shape.Left;
            }

            newSlide.Delete();
        }

        private static void TransferEffects(List<int> effOrder, Shape newGroupedShape,
                                            PowerPointSlide curSlide, PowerPointSlide newSlide)
        {
            foreach (int curo in effOrder)
            {
                Effect eff = newSlide.TimeLine.MainSequence[1];
                eff.Shape = newGroupedShape;

                if (curSlide.TimeLine.MainSequence.Count == 0)
                {
                    Shape tempShape = curSlide.Shapes.AddLine(0, 0, 1, 1);
                    Effect tempEff = curSlide.TimeLine.MainSequence.AddEffect(tempShape, MsoAnimEffect.msoAnimEffectAppear);
                    eff.MoveAfter(tempEff);
                    tempEff.Delete();
                }
                else if (curSlide.TimeLine.MainSequence.Count + 1 < curo)
                {
                    // out of range, assumed to be last
                    eff.MoveAfter(curSlide.TimeLine.MainSequence[curSlide.TimeLine.MainSequence.Count]);
                }
                else if (curo == 1)
                {
                    // first item!
                    eff.MoveBefore(curSlide.TimeLine.MainSequence[1]);
                }
                else
                {
                    eff.MoveAfter(curSlide.TimeLine.MainSequence[curo - 1]);
                }
            }
        }

        private static ShapeRange RemovePlaceholders(PowerPointSlide slide, ShapeRange shapes)
        {
            List<Shape> newShapeList = new List<Shape>();
            foreach (Shape shape in shapes)
            {
                if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                {
                    newShapeList.Add(shape);
                }
            }
            return slide.ToShapeRange(newShapeList);
        }
    }
}
