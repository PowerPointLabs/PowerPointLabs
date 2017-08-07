using System.Collections.Generic;

using PowerPointLabs.NarrationsLab.Views;

namespace PowerPointLabs.NarrationsLab
{
    internal static class NarrationsLabSettings
    {
        public static List<string> VoiceNameList = null;
        public static int VoiceSelectedIndex = 0;

        public static bool IsPreviewEnabled = false;

        public static void ShowSettingsDialog()
        {
            NarrationsLabSettingsDialogBox dialog = new NarrationsLabSettingsDialogBox(
                VoiceSelectedIndex,
                VoiceNameList,
                IsPreviewEnabled);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        private static void OnSettingsDialogConfirmed(string voiceName, bool isPreviewCurrentSlide)
        {
            IsPreviewEnabled = isPreviewCurrentSlide;

            if (!string.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                VoiceSelectedIndex = VoiceNameList.IndexOf(voiceName);
            }
        }
    }
}
