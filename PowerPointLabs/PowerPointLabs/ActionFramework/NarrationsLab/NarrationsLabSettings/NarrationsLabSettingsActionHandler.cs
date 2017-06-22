using System;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Animationlab
{
    [ExportActionRibbonId(TextCollection.NarrationsLabSettingsTag)]
    class NarrationsLabSettingsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var dialog = new NarrationsLabDialogBox(this.GetRibbonUi()._voiceSelected,
                this.GetRibbonUi()._voiceNames, this.GetRibbonUi()._previewCurrentSlide);
            dialog.SettingsHandler += NarrationsLabSettingsChanged;
            dialog.ShowDialog();
        }

        private void NarrationsLabSettingsChanged(String voiceName, bool previewCurrentSlide)
        {
            this.GetRibbonUi()._previewCurrentSlide = previewCurrentSlide;
            if (!String.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                this.GetRibbonUi()._voiceSelected = this.GetRibbonUi()._voiceNames.IndexOf(voiceName);
            }
        }
    }
}
