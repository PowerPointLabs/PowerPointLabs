using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("pasteToFillSlide")]
    class PasteLabFillActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId, bool isClipboardEmpty)
        {
            var curPresentation = this.GetCurrentPresentation();
            PowerPointLabs.PasteLab.PasteLabMain.PasteToFillSlide(this.GetCurrentSlide(), isClipboardEmpty, curPresentation.SlideWidth, curPresentation.SlideHeight);
        }
    }
}
