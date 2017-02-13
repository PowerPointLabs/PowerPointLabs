using System.Windows;

using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PasteLab
{
    public class PasteLabMain
    {
        public static void PasteToFillSlide(Models.PowerPointSlide slide, float width, float height)
        {
            if (IsClipboardEmpty())
            {
                return;
            }

            PowerPoint.ShapeRange pastedObject = slide.Shapes.Paste();
            for (int i = 1; i <= pastedObject.Count; i++)
            {
                var shape = new PPShape(pastedObject[i]);
                shape.AbsoluteHeight = height;
                shape.AbsoluteWidth = width;
                shape.VisualTop = 0;
                shape.VisualLeft = 0;
            }
        }

        internal static bool IsClipboardEmpty()
        {
            return Clipboard.GetDataObject() == null;
        }
    }
}
