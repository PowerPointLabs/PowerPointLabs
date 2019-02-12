using System.Collections.Generic;

using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.Views;

namespace PowerPointLabs.NarrationsLab
{
    internal static class NarrationsLabSettings
    {
        public static List<string> VoiceNameList = null;
        public static int VoiceSelectedIndex = 0;

        public static bool IsPreviewEnabled = false;
        public static AzureVoice azureVoice;

        public static void ShowSettingsDialog()
        {
            NarrationsLabSettingsDialogBox dialog = NarrationsLabSettingsDialogBox.GetInstance();
            NarrationsLabMainSettingsPage.GetInstance().SetNarrationsLabMainSettings(
                VoiceSelectedIndex,
                azureVoice,
                VoiceNameList,
                NotesToAudio.IsAzureVoiceSelected,
                IsPreviewEnabled);
            NarrationsLabMainSettingsPage.GetInstance().DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        private static void OnSettingsDialogConfirmed(string voiceName, AzureVoice voice, bool isAzureVoiceSelected, bool isPreviewCurrentSlide)
        {
            IsPreviewEnabled = isPreviewCurrentSlide;

            if (!string.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                VoiceSelectedIndex = VoiceNameList.IndexOf(voiceName);
            }
            else if (voice != null)
            {
                azureVoice = voice;
                NotesToAudio.SetDefaultVoice(azureVoice.voiceName, azureVoice);
            }
            NotesToAudio.IsAzureVoiceSelected = isAzureVoiceSelected;
        }
    }
}
