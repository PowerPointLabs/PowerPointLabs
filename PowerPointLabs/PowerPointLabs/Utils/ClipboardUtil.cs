using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

namespace PowerPointLabs.Utils
{
#pragma warning disable 0618
    internal static class ClipboardUtil
    {
        public const int ClipboardRestoreSuccess = 1;

        #region API

        public static bool IsClipboardEmpty()
        {
            return PPLClipboard.Instance.IsEmpty();
        }

        /// <summary>
        /// This method assumes that there is valid data on the clipboard. DO NOT lock the clipboard.
        /// </summary>
        public static ShapeRange PasteShapesFromClipboard(PowerPointPresentation pres, PowerPointSlide slide)
        {
            try
            {
                ShapeRange shapes = null;
                try
                {
                    shapes = PasteWithCorrectSlideCheck(slide);
                    // Try to get enumerator and sees if it throws an error
                    // Will throw error if its a web picture
                    shapes.GetEnumerator();
                    return shapes;
                }
                catch (COMException e)
                {
                    Logger.LogException(e, "PasteShapesFromClipboard");
                    // Delete previously pasted "shapes" because it is not a valid shape
                    if (shapes != null)
                    {
                        shapes[1].Delete();
                    }
                    ShapeRange picture = TryPastingAsPNG(slide);
                    if (picture == null)
                    {
                        picture = TryPastingAsBitmap(slide);
                    }
                    if (picture == null)
                    {
                        picture = TryPastingOntoView(pres, slide);
                    }
                    return picture;
                }
            }
            catch (COMException e)
            {
                // May be thrown if there is placeholder shape in clipboard
                Logger.LogException(e, "PasteShapesFromClipboard");
                return null;
            }
        }

        /// <summary>
        /// To avoid changing the clipboard during a copy/cut and paste action. 
        /// One solution for this is to save clipboard into a temp slide and revert clipboard afterwards.
        /// </summary>
        public static TResult RestoreClipboardAfterAction<TResult>(System.Func<TResult> action, PowerPointPresentation pres, PowerPointSlide origSlide)
        {
            TResult result;
            if (!IsClipboardEmpty())
            {
                // Save clipboard onto a temp slide
                PowerPointSlide tempClipboardSlide;
                ShapeRange tempClipboardShapes;
                SlideRange tempPastedSlide;
                SaveClipboard(pres, origSlide, out tempClipboardSlide, out tempClipboardShapes, out tempPastedSlide);

                result = action();

                RestoreClipboard(tempClipboardShapes, tempPastedSlide);
                if (tempClipboardSlide != null)
                {
                    tempClipboardSlide.Delete();
                }
            }
            else
            {
                // Clipboard is empty, we can just run the action function
                result = action();
            }
            return result;
        }

        private static void SaveClipboard(PowerPointPresentation pres, PowerPointSlide origSlide, out PowerPointSlide tempClipboardSlide, out ShapeRange tempClipboardShapes, out SlideRange tempPastedSlide)
        {
            Logger.Log("RestoreClipboardAfterAction: Trying to paste as slide.", ActionFramework.Common.Logger.LogType.Info);
            ClipboardUtilData data = PPLClipboard.Instance.LockAndRelease(() => SaveClipboardUnsafe(pres, origSlide));
            tempClipboardSlide = data.tempClipboardSlide;
            tempClipboardShapes = data.tempClipboardShapes;
            tempPastedSlide = data.tempPastedSlide;
        }

        private static ClipboardUtilData SaveClipboardUnsafe(PowerPointPresentation pres, PowerPointSlide origSlide)
        {
            PowerPointSlide tempClipboardSlide = null;
            ShapeRange tempClipboardShapes = null;
            SlideRange tempPastedSlide = null;

            tempPastedSlide = TryPastingAsSlide(pres, origSlide);

            if (tempPastedSlide == null)
            {
                tempClipboardSlide = pres.AddSlide();
                Logger.Log("RestoreClipboardAfterAction: Trying to paste as text.", ActionFramework.Common.Logger.LogType.Info);
                tempClipboardShapes = TryPastingAsText(tempClipboardSlide);
            }

            if (CheckIfPastingFailed(tempPastedSlide, tempClipboardShapes))
            {
                Logger.Log("RestoreClipboardAfterAction: Trying to paste as shape.", ActionFramework.Common.Logger.LogType.Info);
                tempClipboardShapes = TryPastingAsShape(tempClipboardSlide);
            }

            if (CheckIfPastingFailed(tempPastedSlide, tempClipboardShapes))
            {
                Logger.Log("RestoreClipboardAfterAction: Trying to paste as PNG picture", ActionFramework.Common.Logger.LogType.Info);
                tempClipboardShapes = TryPastingAsPNG(tempClipboardSlide);
            }

            if (CheckIfPastingFailed(tempPastedSlide, tempClipboardShapes))
            {
                Logger.Log("RestoreClipboardAfterAction: Trying to paste as bitmap picture", ActionFramework.Common.Logger.LogType.Info);
                tempClipboardShapes = TryPastingAsBitmap(tempClipboardSlide);
            }

            if (CheckIfPastingFailed(tempPastedSlide, tempClipboardShapes))
            {
                Logger.Log("RestoreClipboardAfterAction: Trying to paste onto view", ActionFramework.Common.Logger.LogType.Info);
                tempClipboardShapes = TryPastingOntoView(pres, tempClipboardSlide, origSlide);
            }
            return new ClipboardUtilData()
            {
                tempClipboardSlide = tempClipboardSlide,
                tempClipboardShapes = tempClipboardShapes,
                tempPastedSlide = tempPastedSlide
            };
        }

        private static bool CheckIfPastingFailed(SlideRange slide, ShapeRange shapes)
        {
            return (slide == null && (shapes == null || shapes.Count < 1));
        }

        #endregion

        /// <summary>
        /// Tries to restore clipboard with provided SlideRange first, then ShapeRange then finally Shape. 
        /// Note that clipboard cannot be restored if last copied item was a placeholder (for now)
        /// </summary>
        /// <returns>True if successfully restored</returns>
        private static void RestoreClipboard(ShapeRange shapes = null, SlideRange slides = null) 
        {
            try
            {
                PPLClipboard.Instance.LockAndRelease(() =>
                {
                    PPLClipboard.Instance.RestoreClipboard();
                    if (slides != null)
                    {
                        slides.Copy();
                        slides.Delete();
                    }
                    else if (shapes != null && shapes.Count >= 1)
                    {
                        shapes.Copy();
                        shapes.Delete();
                    }
                });
            }
            catch (COMException e) 
            {
                // May be thrown when trying to copy
                Logger.LogException(e, "RestoreClipboard");
            }
        }

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

        private static SlideRange TryPastingAsSlide(PowerPointPresentation pres, PowerPointSlide origSlide)
        {
            try
            {
                // try pasting as slide
                SlideRange slides = pres.PasteSlide();
                // Ensure that the view is at the original slide
                pres.GotoSlide(origSlide.Index);
                return (slides.Count >= 1) ? slides : null;
            }
            catch (COMException e)
            {
                // May be thrown if clipboard is not a slide
                Logger.LogException(e, "TryPastingAsSlide");
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
                Logger.LogException(e, "TryPastingAsText");
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
        private static ShapeRange TryPastingAsPNG(PowerPointSlide slide)
        {
            try
            {
                // try pasting as PNG picture to preserve transparency
                return PasteWithCorrectSlideCheck(slide, true, PpPasteDataType.ppPastePNG);
            }
            catch (COMException e)
            {
                // May be thrown if clipboard is not a PNG picture
                Logger.LogException(e, "TryPastingAsPNG");
                return null;
            }
        }

        private static ShapeRange TryPastingAsBitmap(PowerPointSlide slide)
        {
            try
            {
                // try pasting as general bitmap picture
                return PasteWithCorrectSlideCheck(slide, true, PpPasteDataType.ppPasteBitmap);
            }
            catch (COMException e)
            {
                // May be thrown if clipboard is not a picture
                Logger.LogException(e, "TryPastingAsBitmap");
                return null;
            }
        }

        /// <summary>
        /// Pastes clipboard content into new temp slide using the DocumentWindow's View.Paste()
        /// Though this paste will work for most clipboard objects (even web pictures), it could possibly change the undo history.
        /// </summary>
        private static ShapeRange TryPastingOntoView(PowerPointPresentation pres, PowerPointSlide tempSlide, PowerPointSlide origSlide = null)
        {
            try
            {
                // Utilises deprecated Globals class as ClipboardUtil does not utilise ActionFramework
                DocumentWindow workingWindow = Globals.ThisAddIn.Application.ActiveWindow;
                pres.GotoSlide(tempSlide.Index);
                int origShapesCount = tempSlide.Shapes.Count;

                // Note: This will change the undo history
                workingWindow.View.Paste();
                if (origSlide != null)
                {
                    pres.GotoSlide(origSlide.Index);
                }

                int finalShapesCount = tempSlide.Shapes.Count;
                if (finalShapesCount > origShapesCount) 
                {
                    int newShapesCount = finalShapesCount - origShapesCount;
                    int[] shapesToGet = new int[newShapesCount];
                    for (int i = 0; i < shapesToGet.Length; i++)
                    {
                        shapesToGet[i] = origShapesCount + i + 1;
                    }
                    return tempSlide.Shapes.Range(shapesToGet);
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
