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
            using (var synthesizer = new SpeechSynthesizer())
            {
                var installedVoices = synthesizer.GetInstalledVoices();
                var voices = installedVoices.Where(voice => voice.Enabled);
                return voices.Select(voice => voice.VoiceInfo.Name);
            }
        }

        public static void SaveStringToWaveFiles(string notesText, string folderPath, string fileNameFormat)
        {
            var taggedNotes = new TaggedText(notesText);
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
            var builder = GetPromptForText(textToSave);
            PromptToAudio.SaveAsWav(builder, filePath);
        }

        public static void SpeakString(String textToSpeak)
        {
            if (String.IsNullOrWhiteSpace(textToSpeak))
            {
                return;
            }

            var builder = GetPromptForText(textToSpeak);
            PromptToAudio.Speak(builder);
        }

        private static PromptBuilder GetPromptForText(string textToConvert)
        {
            var taggedText = new TaggedText(textToConvert);
            PromptBuilder builder = taggedText.ToPromptBuilder(DefaultVoiceName);
            return builder;
        }
    }
}