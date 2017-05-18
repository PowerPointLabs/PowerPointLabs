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

            // Limitation: Clipboard's shape positions will not be preserved. Unable to find a good fix.
            if (Graphics.IsClipboardEmpty())
            {
                Logger.Log(ribbonId + " failed. Clipboard is empty.");
                return;
            }
            IDataObject clipboardData = Clipboard.GetDataObject();

            ShapeRange pastingShapes = null;
            if (IsSelectionShapes(selection))
            {
                ShapeRange selectedShapes = selection.ShapeRange;
                if (selection.HasChildShapeRange)
                {
                    selectedShapes = selection.ChildShapeRange;
                }
                pastingShapes = slide.Shapes.Paste();
                selectedShapes.Select();
                
            }
            else
            {
                pastingShapes = slide.Shapes.Paste();
            }

            ExecutePasteAction(ribbonId, presentation, slide, selection, pastingShapes);
            Clipboard.SetDataObject(clipboardData);

            CleanupPasteAction();
        }

        protected abstract void ExecutePasteAction(string ribbonId, PowerPointPresentation presentation, PowerPointSlide slide,
                                                    Selection selection, ShapeRange pastingShapes);

        protected virtual void CleanupPasteAction() { }
    }
}
