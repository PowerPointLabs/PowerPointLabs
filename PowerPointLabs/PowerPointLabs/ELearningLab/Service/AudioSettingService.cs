using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Views;

namespace PowerPointLabs.ELearningLab.Service
{
    internal static class AudioSettingService
    {
        public const int AudioMainSettingsPageHeight = 195;
        public const int AudioPreviewPageHeight = 300;

        public static bool IsPreviewEnabled = false;

        public static VoiceType selectedVoiceType = VoiceType.ComputerVoice;
        public static IVoice selectedVoice = ComputerVoiceRuntimeService.Voices.ElementAtOrDefault(0);

        public static void ShowSettingsDialog()
        {
            AudioSettingsDialogWindow dialog = AudioSettingsDialogWindow.GetInstance();
            AudioMainSettingsPage.GetInstance().SetAudioMainSettings(
                selectedVoiceType,
                selectedVoice,
                IsPreviewEnabled);
            AudioMainSettingsPage.GetInstance().DialogConfirmedHandler += OnSettingsDialogConfirmed;
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
