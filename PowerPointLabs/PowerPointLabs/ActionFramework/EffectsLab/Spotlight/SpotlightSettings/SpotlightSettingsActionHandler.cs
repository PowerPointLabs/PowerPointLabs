using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(EffectsLabText.SpotlightSettingsTag)]
    class SpotlightSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            EffectsLabSettings.ShowSpotlightSettingsDialog();
        }
    }
}
