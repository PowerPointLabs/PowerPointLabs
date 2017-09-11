using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

namespace PowerPointLabs.Utils
{
    internal static class ClipboardUtil
    {
        #region API

        public static bool IsClipboardEmpty()
        {
            IDataObject clipboardData = Clipboard.GetDataObject();
            return clipboardData == null || clipboardData.GetFormats().Length == 0;
        }

        public static ShapeRange PasteShapesFromClipboard(PowerPointSlide slide)
        {
            try
            {
                // Note: Some copied objects are pasted on currentSlide rather than the desired slide (e.g. jpg from desktop),
                // so we must check whether it is pasted correctly, else we cut-and-paste it into the correct slide.

                int initialSlideShapesCount = slide.Shapes.Count;
                ShapeRange pastedShapes = slide.Shapes.Paste();

                int finalSlideShapesCount = slide.Shapes.Count;
                if (pastedShapes.Count >= 1 && finalSlideShapesCount == initialSlideShapesCount)
                {
                    pastedShapes.Cut();
                    pastedShapes = slide.Shapes.Paste();
                }

                return pastedShapes;
            }
            catch (COMException e)
            {
                // May be thrown if there is placeholder shape in clipboard
                Logger.LogException(e, "PasteShapeFromClipboard");
                return null;
            }
        }

        #endregion
    }
}
