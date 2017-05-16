using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("pasteToOriginalPosition")]
    class PasteLabOriginalPositionActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId, bool isClipboardEmpty)
        {
            var presentation = this.GetCurrentPresentation();
            var slide = this.GetCurrentSlide();

            PowerPointLabs.PasteLab.PasteLabMain.PasteToOriginalPosition(presentation, slide, isClipboardEmpty);
        }
    }
}