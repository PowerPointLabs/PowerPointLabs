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
            var dialog = new AnimationLabSettingsDialogBox(this.GetRibbonUi().DefaultDuration, this.GetRibbonUi().FrameAnimationChecked);
            dialog.SettingsHandler += AnimationPropertiesEdited;
            dialog.ShowDialog();
        }

        private void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            this.GetRibbonUi().DefaultDuration = newDuration;
            this.GetRibbonUi().FrameAnimationChecked = newFrameChecked;
            AnimateInSlide.animationDuration = newDuration;
            AnimateInSlide.frameAnimationChecked = newFrameChecked;
            AutoAnimate.duration = newDuration;
            AutoAnimate.frameAnimationChecked = newFrameChecked;
        }
    }
}
