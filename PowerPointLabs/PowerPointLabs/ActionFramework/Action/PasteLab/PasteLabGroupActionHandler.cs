using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("pasteIntoGroup")]
    class PasteLabGroupActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId, bool isClipboardEmpty)
        {
            var presentation = this.GetCurrentPresentation();
            var slide = this.GetCurrentSlide();
            var selection = this.GetCurrentSelection();

            if (!IsSelectionShapes(selection))
            {
                Logger.Log("PasteIntoGroup failed. No valid shape is selected.");
                return;
            }

            if (isClipboardEmpty)
            {
                Logger.Log("PasteIntoGroup failed. Clipboard is empty.");
                return;
            }

            ShapeRange pastingShapes = slide.Shapes.Paste();
            PowerPointLabs.PasteLab.PasteLabMain.PutIntoGroup(presentation, slide, selection.ShapeRange, pastingShapes);
        }
    }
}