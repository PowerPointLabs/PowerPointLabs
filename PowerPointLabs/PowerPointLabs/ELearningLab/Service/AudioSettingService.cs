using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ELearningLab.Views;

namespace PowerPointLabs.ELearningLab.Service
{
#pragma warning disable 618
    internal static class AudioSettingService
    {
        public const int AudioMainSettingsPageHeight = 195;
        public static double AudioPreviewPageHeight = 300;

        public static bool IsPreviewEnabled = false;

        public static VoiceType selectedVoiceType = VoiceType.ComputerVoice;
        public static IVoice selectedVoice = ComputerVoiceRuntimeService.Voices.ElementAtOrDefault(0);
        public static ObservableCollection<IVoice> preferredVoices
             = new ObservableCollection<IVoice>();

        public static void ShowSettingsDialog()
        {
            CustomTaskPane eLearningTaskpane = ActionFrameworkExtensions.GetTaskPane(typeof(ELearningLabTaskpane));
            AudioSettingsDialogWindow dialog = new AudioSettingsDialogWindow(AudioSettingsPage.MainSettingsPage);
            AudioMainSettingsPage page = dialog.MainPage as AudioMainSettingsPage;
            page.SetAudioMainSettings(
                selectedVoiceType,
                selectedVoice,
                IsPreviewEnabled);
            page.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            if (eLearningTaskpane == null)
            {
                dialog.ShowDialog();
                return;
            }
            ELearningLabTaskpane taskpane = eLearningTaskpane.Control as ELearningLabTaskpane;
            page.DefaultVoiceChangedHandler +=
                taskpane.ELearningLabMainPanel.RefreshVoiceLabelOnAudioSettingChanged;
            page.IsDefaultVoiceChangedHandlerAssigned = true;
            dialog.ShowDialog();
        }

        private static void OnSettingsDialogConfirmed(VoiceType selectedVoiceType, IVoice selectedVoice, bool isPreviewCurrentSlide)
        {
            IsPreviewEnabled = isPreviewCurrentSlide;

            AudioSettingService.selectedVoiceType = selectedVoiceType;
            AudioSettingService.selectedVoice = selectedVoice;
        }
    }
}
