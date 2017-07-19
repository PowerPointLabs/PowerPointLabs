using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportActionRibbonId(TextCollection.RemoveHighlightTag)]
    class RemoveHighlightActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            RemoveHighlighting.RemoveHighlight(this.GetCurrentSlide());
        }
    }
}
