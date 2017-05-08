﻿using System;
using System.Collections.Generic;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PasteLab
{
    public class PasteLabMain
    {
        public static void PasteToFillSlide(Models.PowerPointSlide slide, bool clipboardIsEmpty, float width, float height)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToFillSlide encountered empty clipboard");
                return;
            }

            PowerPoint.ShapeRange pastedObject = slide.Shapes.Paste();

            Logger.Log(string.Format("PasteToFillSlide: {0} objects pasted", pastedObject.Count));

            for (int i = 1; i <= pastedObject.Count; i++)
            {
                var shape = new PPShape(pastedObject[i]);
                shape.AbsoluteHeight = height;

                if (shape.AbsoluteWidth < width)
                {
                    shape.AbsoluteWidth = width;
                }

                shape.VisualCenter = new System.Drawing.PointF(width / 2, height / 2);
            }
        }

        public static void PasteAndReplace(Models.PowerPointPresentation presentation, Models.PowerPointSlide slide,
                                           bool clipboardIsEmpty, PowerPoint.Selection selection)
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

            PowerPoint.Shape newShape = slide.Shapes.Paste()[1];
            newShape.Left = shapeToReplace.Left;
            newShape.Top = shapeToReplace.Top;

            foreach (PowerPoint.Effect eff in slide.TimeLine.MainSequence)
            {
                if (eff.Shape == shapeToReplace)
                {
                    PowerPoint.Effect newEff = slide.TimeLine.MainSequence.Clone(eff);
                    newEff.Shape = newShape;
                    eff.Delete();
                }
            }

            shapeToReplace.PickUp();
            newShape.Apply();

            Logger.Log(string.Format("PasteAndReplace: Replaced {0} with {1}", shapeToReplace.Name, newShape.Name));
            shapeToReplace.Delete();
        }

        public static PowerPoint.Shape PasteIntoGroup(Models.PowerPointPresentation presentation, Models.PowerPointSlide slide,
                                                           bool clipboardIsEmpty, PowerPoint.Selection selection)
        {
            var selectedShapes = selection.ShapeRange;
            var pastedShapes = slide.Shapes.Paste();

            var tempSlide = presentation.AddSlide();
            selectedShapes.Copy();
            tempSlide.Shapes.Paste();
            pastedShapes.Copy();    // revert the clipboard state

            List<int> transferEffects = new List<int>();
            foreach (PowerPoint.Effect effect in slide.TimeLine.MainSequence)
            {
                if (effect.Shape.Equals(selectedShapes[1]))
                {
                    transferEffects.Add(effect.Index);
                }
            }
            List<String> transferShapeNames = new List<String>();
            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                transferShapeNames.Add(shape.Name);
            }
            foreach (PowerPoint.Shape shape in pastedShapes)
            {
                transferShapeNames.Add(shape.Name);
            }
            PowerPoint.ShapeRange transferShapes = slide.Shapes.Range(transferShapeNames.ToArray());

            float selectionLeft = selectedShapes[1].Left;
            float selectionTop = selectedShapes[1].Top;
            float selectionWidth = selectedShapes[1].Width;
            float selectionHeight = selectedShapes[1].Height;
            if (selectedShapes.Count > 1)
            {
                var selectionGroup = selectedShapes.Group();
                selectionLeft = selectionGroup.Left;
                selectionTop = selectionGroup.Top;
                selectionWidth = selectionGroup.Width;
                selectionHeight = selectionGroup.Height;
                selectedShapes.Ungroup();
            }

            // Paste at center of the selection
            if (pastedShapes.Count > 1)
            {
                var pastedGroup = pastedShapes.Group();
                pastedGroup.Left = selectionLeft + (selectionWidth - pastedGroup.Width) / 2;
                pastedGroup.Top = selectionTop + (selectionHeight - pastedGroup.Height) / 2;
                pastedShapes.Ungroup();
            }
            else
            {
                pastedShapes[1].Left = selectionLeft + (selectionWidth - pastedShapes[1].Width) / 2;
                pastedShapes[1].Top = selectionTop + (selectionHeight - pastedShapes[1].Height) / 2;
            }

            PowerPoint.Shape transferShapesGroup = transferShapes.Group();
            TransferEffects(transferEffects, transferShapesGroup, slide, tempSlide);

            tempSlide.Delete();
            return transferShapesGroup;
        }

        public static void GroupSelectedShapes(Models.PowerPointPresentation presentation, Models.PowerPointSlide slide,
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

            foreach (PowerPoint.Effect eff in slide.TimeLine.MainSequence)
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

        public static PowerPoint.ShapeRange PasteToPosition(Models.PowerPointSlide slide, bool clipboardIsEmpty, float xPosition, float yPosition)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToPosition encountered an empty clipboard");
                return null;
            }

            PowerPoint.ShapeRange pastedShapes = slide.Shapes.Paste();
            foreach (PowerPoint.Shape shape in pastedShapes)
            {
                shape.Left = xPosition;
                shape.Top = yPosition;
            }

            return pastedShapes;
        }

        public static void PasteToOriginalPosition(Models.PowerPointPresentation presentation,
                                                   Models.PowerPointSlide slide, bool clipboardIsEmpty)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToOriginalPosition encountered an empty clipboard");
                return;
            }

            // This is identical to PowerPoint's native paste function.
            slide.Shapes.Paste();
        }

        private static void TransferEffects(List<int> effOrder, PowerPoint.Shape newGroupedShape,
                                            Models.PowerPointSlide curSlide, Models.PowerPointSlide newSlide)
        {
            foreach (int curo in effOrder)
            {
                PowerPoint.Effect eff = newSlide.TimeLine.MainSequence[1];
                eff.Shape = newGroupedShape;

                if (curSlide.TimeLine.MainSequence.Count == 0)
                {
                    PowerPoint.Shape tempShape = curSlide.Shapes.AddLine(0, 0, 1, 1);
                    PowerPoint.Effect tempEff = curSlide.TimeLine.MainSequence.AddEffect(tempShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear);
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
    }
}
