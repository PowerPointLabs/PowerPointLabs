using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using PowerPointLabs.Models;

namespace PowerPointLabs.SpeechEngine
{
    static class TextToSpeech
    {
        public static String DefaultVoiceName;

        public static IEnumerable<string> GetVoices()
        {
            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                System.Collections.ObjectModel.ReadOnlyCollection<InstalledVoice> installedVoices = synthesizer.GetInstalledVoices();
                IEnumerable<InstalledVoice> voices = installedVoices.Where(voice => voice.Enabled);
                return voices.Select(voice => voice.VoiceInfo.Name);
            }
        }

        public static void SaveStringToWaveFiles(string notesText, string folderPath, string fileNameFormat)
        {
            TaggedText taggedNotes = new TaggedText(notesText);
            List<String> stringsToSave = taggedNotes.SplitByClicks();
            //MD5 md5 = MD5.Create();

            for (int i = 0; i < stringsToSave.Count; i++)
            {
                String textToSave = stringsToSave[i];
                String baseFileName = String.Format(fileNameFormat, i + 1);

                // The first item will autoplay; everything else is triggered by a click.
                String fileName = i > 0 ? baseFileName + " (OnClick)" : baseFileName;

                String filePath = folderPath + "\\" + fileName + ".wav";

                SaveStringToWaveFile(textToSave, filePath);
            }
        }

        public static void SaveStringToWaveFile(String textToSave, String filePath)
        {
            PromptBuilder builder = GetPromptForText(textToSave);
            PromptToAudio.SaveAsWav(builder, filePath);
        }

        public static void SpeakString(String textToSpeak)
        {
            if (String.IsNullOrWhiteSpace(textToSpeak))
            {
                return;
            }

            PromptBuilder builder = GetPromptForText(textToSpeak);
            PromptToAudio.Speak(builder);
        }

        private static PromptBuilder GetPromptForText(string textToConvert)
        {
            TaggedText taggedText = new TaggedText(textToConvert);
            PromptBuilder builder = taggedText.ToPromptBuilder(DefaultVoiceName);
            return builder;
        }
    }
}