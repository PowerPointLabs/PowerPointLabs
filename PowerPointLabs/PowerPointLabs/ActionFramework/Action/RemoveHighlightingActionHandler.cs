using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("RemoveHighlightButton")]
    class RemoveHighlightingHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.GetApplication().StartNewUndoEntry();
            RemoveHighlighting.RemoveAllHighlighting();
        }
    }
}
