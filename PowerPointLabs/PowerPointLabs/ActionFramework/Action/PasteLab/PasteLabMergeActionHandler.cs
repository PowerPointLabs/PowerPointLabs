using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    [ExportActionRibbonId("mergeSelection")]
    class PasteLabMergeActionHandler : PasteLabActionHandler
    {
        protected override void ExecutePasteAction(string ribbonId)
        {
            var presentation = this.GetCurrentPresentation();
            var slide = this.GetCurrentSlide();
            var selection = this.GetCurrentSelection();

            PowerPointLabs.PasteLab.PasteLabMain.GroupSelectedShapes(presentation, slide, selection);
        }
    }
}