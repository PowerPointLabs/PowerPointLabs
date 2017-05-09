using System;
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

        public static PowerPoint.ShapeRange PasteIntoGroup(Models.PowerPointPresentation presentation, Models.PowerPointSlide slide,
                                                           bool clipboardIsEmpty, PowerPoint.Selection selection)
        {
            var newSlide = presentation.AddSlide();
            var selectedShapes = selection.ShapeRange;

            PowerPoint.ShapeRange pastedShapes = slide.Shapes.Paste();

            selection.Copy();
            newSlide.Shapes.Paste();
            pastedShapes.Copy();

            List<int> order = new List<int>();

            foreach (PowerPoint.Effect eff in slide.TimeLine.MainSequence)
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

            List<int> effectsOrder = new List<int>();
            foreach (PowerPoint.Effect eff in slide.TimeLine.MainSequence)
            {
                if (eff.Shape.Equals(selectedShapes[1]))
                {
                    effectsOrder.Add(eff.Index);
                }
            }

            PowerPoint.Shape newGroupedShape = selectedShapes.Group();
            TransferEffects(effectsOrder, newGroupedShape, slide, newSlide);
            newSlide.Delete();
        }

        public static void PasteToPosition(Models.PowerPointSlide slide, bool clipboardIsEmpty, float xPosition, float yPosition)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToPosition encountered an empty clipboard");
                return;
            }

            var newShapeRange = slide.Shapes.Paste();

            foreach (PowerPoint.Shape shape in newShapeRange)
            {
                shape.Left = xPosition;
                shape.Top = yPosition;

                Logger.Log(string.Format("PasteToPosition: Pasted {0} at ({1}, {2})", shape.Name, shape.Left, shape.Top));
            }
        }

        public static void PasteToOriginalPosition(Models.PowerPointPresentation presentation,
                                                   Models.PowerPointSlide slide, bool clipboardIsEmpty)
        {
            if (clipboardIsEmpty)
            {
                Logger.Log("PasteToOriginalPosition encountered an empty clipboard");
                return;
            }

            // Needs new slide, otherwise there will be a slight offset when pasting
            var newSlide = presentation.AddSlide();

            PowerPoint.ShapeRange correctShapes = newSlide.Shapes.Paste();

            foreach (PowerPoint.Shape shape in correctShapes)
            {
                shape.Copy();
                PowerPoint.Shape pastedShape = slide.Shapes.Paste()[1];
                pastedShape.Top = shape.Top;
                pastedShape.Left = shape.Left;
            }

            newSlide.Delete();
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
