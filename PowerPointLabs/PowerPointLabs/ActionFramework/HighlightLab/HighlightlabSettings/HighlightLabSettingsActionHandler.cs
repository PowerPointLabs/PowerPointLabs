using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.HighlightLab;

namespace PowerPointLabs.ActionFramework.HighlightLab
{
    [ExportActionRibbonId(TextCollection1.HighlightLabSettingsTag)]
    class HighlightLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            HighlightLabSettings.ShowSettingsDialog();
        }
    }
}
