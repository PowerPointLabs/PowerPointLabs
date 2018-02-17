using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

namespace PowerPointLabs.Utils
{
#pragma warning disable 0618
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
                return PasteWithCorrectSlideCheck(slide);
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
        public static void RestoreClipboardAfterAction(System.Action action, PowerPointPresentation pres, PowerPointSlide origSlide)
        {
            if (!IsClipboardEmpty())
            {
                // Save clipboard onto a temp slide
                PowerPointSlide tempClipboardSlide = pres.AddSlide();
                ShapeRange tempClipboardShapes = null;
                SlideRange tempPastedSlide = null;
                Shape tempClipboardShape = null;

                Logger.Log("RestoreClipboardAfterAction: Trying to paste as slide.", ActionFramework.Common.Logger.LogType.Info);
                tempPastedSlide = TryPastingAsSlide(pres);

                if (tempPastedSlide == null)
                {
                    Logger.Log("RestoreClipboardAfterAction: Trying to paste as text.", ActionFramework.Common.Logger.LogType.Info);
                    tempClipboardShapes = TryPastingAsText(tempClipboardSlide);
                }

                if (tempPastedSlide == null && (tempClipboardShapes == null || tempClipboardShapes.Count < 1))
                {
                    Logger.Log("RestoreClipboardAfterAction: Trying to paste as shape.", ActionFramework.Common.Logger.LogType.Info);
                    tempClipboardShapes = TryPastingAsShape(tempClipboardSlide);
                }

                if (tempPastedSlide == null && (tempClipboardShapes == null || tempClipboardShapes.Count < 1))
                {
                    Logger.Log("RestoreClipboardAfterAction: Trying to paste onto view normally.", ActionFramework.Common.Logger.LogType.Info);
                    tempClipboardShape = TryPastingOntoView(pres, tempClipboardSlide, origSlide);
                }

                action();

                try
                {
                    // Revert clipboard. Note that clipboard cannot be reverted if last copied item was a placeholder (for now)
                    if (tempPastedSlide != null)
                    {
                        tempPastedSlide.Copy();
                        tempPastedSlide.Delete();
                    }
                    else if (tempClipboardShapes != null && tempClipboardShapes.Count >= 1)
                    {
                        tempClipboardShapes.Copy();
                        tempClipboardShapes.Delete();
                    } 
                    else if (tempClipboardShape != null) 
                    {
                        tempClipboardShape.Copy();
                        tempClipboardShape.Delete();
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

        private static ShapeRange PasteWithCorrectSlideCheck(PowerPointSlide slide, bool isPasteSpecial = false, PpPasteDataType pasteType = PpPasteDataType.ppPasteDefault)
        {
            // Note: Some copied objects are pasted on currentSlide rather than the desired slide (e.g. jpg from desktop),
            // so we must check whether it is pasted correctly, else we cut-and-paste it into the correct slide.

            int initialSlideShapesCount = slide.Shapes.Count;
            ShapeRange pastedShapes = null;
            if (!isPasteSpecial)
            {
                pastedShapes = slide.Shapes.Paste();
            }
            else
            {
                pastedShapes = slide.Shapes.PasteSpecial(pasteType);
            }

            int finalSlideShapesCount = slide.Shapes.Count;
            if (pastedShapes.Count >= 1 && finalSlideShapesCount == initialSlideShapesCount)
            {
                pastedShapes.Cut();
                if (!isPasteSpecial)
                {
                    pastedShapes = slide.Shapes.Paste();
                }
                else
                {
                    pastedShapes = slide.Shapes.PasteSpecial(pasteType);
                }
            }

            return pastedShapes;
        }

        private static SlideRange TryPastingAsSlide(PowerPointPresentation pres)
        {
            try
            {
                // try pasting as slide
                SlideRange slides = pres.PasteSlide();
                return (slides.Count >= 1) ? slides : null;
            }
            catch (COMException e)
            {
                // May be thrown if clipboard is not a slide
                Logger.LogException(e, "RestoreClipboardAfterAction: pasting as slide");
                return null;
            }
        }

        private static ShapeRange TryPastingAsText(PowerPointSlide slide)
        {
            try
            {
                // try pasting as text
                return PasteWithCorrectSlideCheck(slide, true, PpPasteDataType.ppPasteText);
            }
            catch (COMException e)
            {
                // May be thrown if clipboard is not text
                Logger.LogException(e, "RestoreClipboardAfterAction: pasting as text");
                return null;
            }
        }

        private static ShapeRange TryPastingAsShape(PowerPointSlide slide) 
        {
            try
            {
                // try pasting as shape
                return PasteWithCorrectSlideCheck(slide, true, PpPasteDataType.ppPasteShape);
            }
            catch (COMException e)
            {
                // May be thrown if clipboard is not a shape
                Logger.LogException(e, "TryPastingAsShape");
                return null;
            }
        }

        /// <summary>
        /// Pastes clipboard content into new temp slide using the DocumentWindow's View.Paste()
        /// Though this paste will work for most clipboard objects (even web pictures), it will change the undo history
        /// </summary>
        private static Shape TryPastingOntoView(PowerPointPresentation pres, PowerPointSlide tempSlide, PowerPointSlide origSlide)
        {
            try
            {
                // Utilises deprecated Globals class as ClipboardUtil does not utilise ActionFramework
                DocumentWindow workingWindow = Globals.ThisAddIn.Application.ActiveWindow;
                // Note: This will change the undo history
                pres.GotoSlide(tempSlide.Index);
                int origShapesCount = tempSlide.Shapes.Count;

                workingWindow.View.Paste();
                pres.GotoSlide(origSlide.Index);
                int finalShapesCount = tempSlide.Shapes.Count;
                if (finalShapesCount > origShapesCount) 
                {
                    return tempSlide.Shapes.Range()[finalShapesCount];
                } 
                else 
                {
                    return null;
                }
            }
            catch (COMException e)
            {
                // May be thrown if cannot be pasted
                Logger.LogException(e, "TryPastingOntoView");
                return null;
            }
        }
    }
}
