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

        /// <summary>
        /// To avoid changing the clipboard during a copy/cut and paste action. 
        /// One solution for this is to save clipboard into a temp slide and revert clipboard afterwards.
        /// </summary>
        public static void RestoreClipboardAfterAction(System.Action action, PowerPointPresentation pres)
        {
            if (!IsClipboardEmpty())
            {
                // Save clipboard onto a temp slide
                PowerPointSlide tempClipboardSlide = pres.AddSlide();
                ShapeRange tempClipboardShapes = PasteShapesFromClipboard(tempClipboardSlide);
                action();

                try
                {
                    // Revert clipboard. Note that clipboard cannot be reverted if last copied item was a placeholder (for now)
                    if (tempClipboardShapes != null && tempClipboardShapes.Count >= 1)
                    {
                        tempClipboardShapes.Copy();
                        tempClipboardShapes.Delete();
                            //System.Runtime.InteropServices.Marshal.ReleaseComObject(tempClipboardShapes);
                    }
                }
                catch (COMException e)
                {
                    // May be thrown when trying to copy if previous clipboard item was not a shape (eg. a slide, certain web pictures)
                    Logger.LogException(e, "RestoreClipboardAfterAction");
                }
                finally
                {
                    tempClipboardSlide.Delete();
                }
            }
            else
            {
                // Clipboard is empty, we can just run the action function
                action();
            }
            
        }
           

        #endregion
    }
}
