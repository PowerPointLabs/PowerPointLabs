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
    }
}
