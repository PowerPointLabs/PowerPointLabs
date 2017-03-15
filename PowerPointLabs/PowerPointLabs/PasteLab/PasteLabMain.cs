using System;
using System.Collections.Generic;

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

        public static void PasteAndReplace(Models.PowerPointSlide slide, bool clipboardIsEmpty, PowerPoint.Selection selection)
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

        public static void PasteIntoGroup(Models.PowerPointPresentation presentation, Models.PowerPointSlide slide,
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
