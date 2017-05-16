using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    abstract class PasteLabActionHandler : ActionHandler
    {
        // Sealed method: Subclasses should override ExecutePasteAction instead
        protected sealed override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            // Store and restore clipboard data:
            // Reason for not using Clipboard.SetDataObject(): it does not preserve position
            var currentSelectedShapes = this.GetCurrentSelection().ShapeRange;
            var tempSlide = this.GetCurrentPresentation().AddSlide(index: this.GetCurrentSlide().Index);
            ShapeRange clipboardItems = tempSlide.Shapes.Paste();
            currentSelectedShapes.Select();
            
            ExecutePasteAction(ribbonId);

            clipboardItems.Copy();
            tempSlide.Delete();
        }

        protected abstract void ExecutePasteAction(string ribbonId);
    }
}
