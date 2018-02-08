using System;
using System.Linq;
using System.Speech.Synthesis;

namespace PowerPointLabs.Tags
{
    public class StartVoiceTag : Tag
    {
        public StartVoiceTag(int start, int end, string contents)
        {
            Start = start;
            End = end;
            Contents = contents;
        }

        public override bool Apply(PromptBuilder builder)
        {
            builder.EndVoice();

            String voiceArgument = ParseTagArgument().ToLower();

            String voiceName = FindFullVoiceName(voiceArgument);

            if (voiceName == null)
            {
                return false;
            }

            builder.StartVoice(voiceName);
            return true;
        }

        public override string PrettyPrint()
        {
            return "";
        }

        private static string FindFullVoiceName(string voiceArgument)
        {
            string voiceName = null;
            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                System.Collections.ObjectModel.ReadOnlyCollection<InstalledVoice> installedVoices = synthesizer.GetInstalledVoices();
                System.Collections.Generic.IEnumerable<InstalledVoice> enabledVoices = installedVoices.Where(voice => voice.Enabled);

                InstalledVoice selectedVoice = enabledVoices.FirstOrDefault(voice => voice.VoiceInfo.Name.ToLowerInvariant().Contains(voiceArgument));
                if (selectedVoice != null)
                {
                    voiceName = selectedVoice.VoiceInfo.Name;
                }
            }
            return voiceName;
        }
    }
}