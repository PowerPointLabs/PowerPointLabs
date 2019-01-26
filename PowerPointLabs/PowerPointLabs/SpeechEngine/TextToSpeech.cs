using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Log;

using PowerPointLabs.Models;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.ViewModel;

namespace PowerPointLabs.SpeechEngine
{
    static class TextToSpeech
    {
        public static String DefaultVoiceName;
        public static AzureVoice humanVoice;

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
                if (NotesToAudio.IsAzureVoiceSelected)
                {
                    SaveStringToWaveFileWithAzureVoice(textToSave, filePath);
                }
                else
                {
                    SaveStringToWaveFile(textToSave, filePath);
                }
            }
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

        private static void SaveStringToWaveFileWithAzureVoice(string textToSave, string filePath)
        {
            string accessToken;
            string textToSpeak = GetHumanSpeakNotesForText(textToSave);
            try
            {
                Authentication auth = Authentication.GetInstance();
                accessToken = auth.GetAccessToken();
                Logger.Log("Token: " + accessToken);
            }
            catch (Exception ex)
            {
                Logger.Log("Failed authentication.");
                Logger.Log(ex.ToString());
                Logger.Log(ex.Message);
                return;
            }
            string requestUri = UserAccount.GetInstance().GetUri();
            if (requestUri == null)
            {
                return;
            }
            var cortana = new Synthesize();

            cortana.OnAudioAvailable += SaveAudioToWaveFile;
            cortana.OnError += ErrorHandler;

            // Reuse Synthesize object to minimize latency
            cortana.Speak(CancellationToken.None, new Synthesize.InputOptions()
            {
                RequestUri = new Uri(requestUri),
                Text = textToSpeak,
                VoiceType = humanVoice.voiceType,
                Locale = humanVoice.Locale,
                VoiceName = humanVoice.voiceName,
                // Service can return audio in different output format.
                OutputFormat = AudioOutputFormat.Riff24Khz16BitMonoPcm,
                AuthorizationToken = "Bearer " + accessToken,
            }, filePath).Wait();
        }

        private static void SaveStringToWaveFile(String textToSave, String filePath)
        {
            PromptBuilder builder = GetPromptForText(textToSave);
            PromptToAudio.SaveAsWav(builder, filePath);
        }

        private static string GetHumanSpeakNotesForText(string textToSave)
        {
            TaggedText taggedText = new TaggedText(textToSave);
            string strToSpeak = taggedText.ToPrettyString();
            return strToSpeak;
        }
        private static PromptBuilder GetPromptForText(string textToConvert)
        {
            TaggedText taggedText = new TaggedText(textToConvert);
            PromptBuilder builder = taggedText.ToPromptBuilder(DefaultVoiceName);
            return builder;
        }

        private static void SaveAudioToWaveFile(object sender, GenericEventArgs<Stream> args)
        {
            //  Logger.Log("saving " + args.FilePath);
            SaveStreamToFile(args.FilePath, args.EventData);
            args.EventData.Dispose();
        }

        private static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];

            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
        private static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            try
            {
                byte[] bytesInStream = ReadFully(stream);
                using (FileStream fs = File.Create(fileFullPath))
                {
                    Console.WriteLine("file created");
                    fs.Write(bytesInStream, 0, bytesInStream.Length);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Error generating audio files. ");
                Debug.WriteLine(e.Message);
            }
        }
        private static void ErrorHandler(object sender, GenericEventArgs<Exception> e)
        {
            Logger.Log("Unable to complete the TTS request: " + e.ToString());
        }
    }
}