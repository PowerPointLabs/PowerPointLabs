using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.AnimationLab;
using PowerPointLabs.AnimationLab.Views;

namespace PowerPointLabs.ActionFramework.Animationlab
{
    [ExportActionRibbonId(TextCollection.AnimationLabSettingsTag)]
    class AnimationLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new AnimationLabSettingsDialogBox(AnimationLabSettings.AnimationDuration, AnimationLabSettings.IsUseFrameAnimation);
            dialog.SettingsHandler += AnimationPropertiesEdited;
            dialog.ShowDialog();
        }

        private void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            AnimationLabSettings.AnimationDuration = newDuration;
            AnimationLabSettings.IsUseFrameAnimation = newFrameChecked;
        }
    }
}
