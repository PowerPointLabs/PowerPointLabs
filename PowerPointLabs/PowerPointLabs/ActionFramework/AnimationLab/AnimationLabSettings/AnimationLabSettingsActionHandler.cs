using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.AnimationLab;

namespace PowerPointLabs.ActionFramework.AnimationLab
{
    [ExportActionRibbonId(TextCollection1.AnimationLabSettingsTag)]
    class AnimationLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            AnimationLabSettings.ShowSettingsDialog();
        }
    }
}
