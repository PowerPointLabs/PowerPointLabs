using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.AnimationLab;

namespace PowerPointLabs.ActionFramework.Animationlab
{
    [ExportActionRibbonId(TextCollection.AnimationLabSettingsTag)]
    class AnimationLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            AnimationLabSettings.ShowSettingsDialog();
        }
    }
}
