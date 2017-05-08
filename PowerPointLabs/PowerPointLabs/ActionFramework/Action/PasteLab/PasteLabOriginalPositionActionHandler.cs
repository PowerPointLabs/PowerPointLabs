using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("pasteToOriginalPosition")]
    class PasteLabOriginalPositionActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var presentation = this.GetCurrentPresentation();
            var slide = this.GetCurrentSlide();
            bool clipboardIsEmpty = (Clipboard.GetDataObject() == null);

            this.StartNewUndoEntry();
            PowerPointLabs.PasteLab.PasteLabMain.PasteToOriginalPosition(presentation, slide, clipboardIsEmpty);
        }
    }
}