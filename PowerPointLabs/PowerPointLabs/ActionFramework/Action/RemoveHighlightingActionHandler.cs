using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("RemoveHighlightButton")]
    class RemoveHighlightingHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.GetApplication().StartNewUndoEntry();
            RemoveHighlighting.RemoveHighlight(this.GetCurrentSlide());
        }
    }
}
