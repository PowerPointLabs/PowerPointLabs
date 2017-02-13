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
            //this.StartNewUndoEntry();
            var curPresentation = this.GetCurrentPresentation();
            PowerPointLabs.PasteLab.PasteLabMain.PasteToFillSlide(this.GetCurrentSlide(), curPresentation.SlideWidth, curPresentation.SlideHeight);
        }
    }
}
