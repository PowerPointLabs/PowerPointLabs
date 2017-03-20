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
    }
}
