using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Animationlab
{
    [ExportActionRibbonId(TextCollection.AnimationLabSettingsTag)]
    class AnimationLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new AnimationLabDialogBox(this.GetRibbonUi().DefaultDuration, this.GetRibbonUi().FrameAnimationChecked);
            dialog.SettingsHandler += AnimationPropertiesEdited;
            dialog.ShowDialog();
        }

        private void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            this.GetRibbonUi().DefaultDuration = newDuration;
            this.GetRibbonUi().FrameAnimationChecked = newFrameChecked;
            AnimateInSlide.defaultDuration = newDuration;
            AnimateInSlide.frameAnimationChecked = newFrameChecked;
            AutoAnimate.defaultDuration = newDuration;
            AutoAnimate.frameAnimationChecked = newFrameChecked;
        }
    }
}
