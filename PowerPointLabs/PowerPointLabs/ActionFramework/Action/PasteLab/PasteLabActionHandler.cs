using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ActionFramework.Util;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    abstract class PasteLabActionHandler : BaseUtilActionHandler
    {
        // Sealed method: Subclasses should override ExecutePasteAction instead
        protected sealed override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            PowerPointPresentation presentation = this.GetCurrentPresentation();
            PowerPointSlide slide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            if (Graphics.IsClipboardEmpty())
            {
                Logger.Log(ribbonId + " failed. Clipboard is empty.");
                return;
            }

            ShapeRange tempClipboardShapes = null;
            PowerPointSlide tempClipboardSlide = presentation.AddSlide(index: slide.Index);
            if (IsSelectionShapes(selection))
            {
                ShapeRange selectedShapes = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                {
                    selectedShapes = selection.ChildShapeRange;
                }
                tempClipboardShapes = tempClipboardSlide.Shapes.Paste();
                selectedShapes.Select();
            }
            else
            {
                tempClipboardShapes = tempClipboardSlide.Shapes.Paste();
            }
            ShapeRange pastingShapes = slide.CopyShapesToSlide(tempClipboardShapes);
            ShapeRange result = ExecutePasteAction(ribbonId, presentation, slide, selection, pastingShapes);
            if (result != null)
            {
                result.Select();
            }

            tempClipboardShapes.Copy();
            tempClipboardShapes.Delete();
            tempClipboardSlide.Delete();
        }

        protected abstract ShapeRange ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                        Selection selection, ShapeRange pastingShapes);
    }
}
