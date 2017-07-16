using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;

namespace PowerPointLabs.ActionFramework.Highlightlab
{
    [ExportActionRibbonId(TextCollection.HighlightLabSettingsTag)]
    class HighlightLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            HighlightLabSettings.ShowSettingsDialog();
        }
    }
}
