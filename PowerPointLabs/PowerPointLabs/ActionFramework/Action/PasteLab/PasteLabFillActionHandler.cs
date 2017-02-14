using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("pasteToFillSlide")]
    class PasteLabFillActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var curPresentation = this.GetCurrentPresentation();
            bool clipboardIsEmpty = (Clipboard.GetDataObject() == null);
            PowerPointLabs.PasteLab.PasteLabMain.PasteToFillSlide(this.GetCurrentSlide(), clipboardIsEmpty, curPresentation.SlideWidth, curPresentation.SlideHeight);
        }
    }
}
