using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

using NAudio.Wave;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.Model;
using PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.SpeechEngine;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Service
{
    public class WatsonRuntimeService
    {
        public static bool IsWatsonAccountPresentAndValid = false;
        public static ObservableCollection<WatsonVoice> Voices { get; set; } = GetWatsonVoices();

        public static bool IsWatsonAccountPresent()
        {
            return !WatsonAccount.GetInstance().IsEmpty();
        }

        public static bool IsValidUserAccount(bool showErrorMessage = true, string errorMessage = "Failed Azure authentication.")
        {
            try
            {
                string _key = WatsonAccount.GetInstance().GetKey();
                if (_key == null || string.IsNullOrEmpty(_key.Trim()))
                {
                    throw new Exception("Empty key value");
                }
                TokenOptions options = new TokenOptions();
                options.IamApiKey = _key;
                string accessToken = new TokenManager(options).GetToken();
                if (string.IsNullOrEmpty(accessToken))
                {
                    throw new Exception("Invalid access key!");
                }
                if (!IsValidEndpoint(_key, WatsonAccount.GetInstance().GetEndpoint()))
                {
                    throw new Exception("Invalid endpoint!");
                }
                Console.WriteLine("Token: {0}\n", accessToken);
            }
            catch
            {
                Console.WriteLine(errorMessage);
                if (showErrorMessage)
                {
                    MessageBox.Show(errorMessage);
                }
                return false;
            }
            return true;
        }

        public static bool IsValidUserAccount(string key, string endpoint, string errorMessage = "Failed Azure authentication.")
        {
            try
            {
                TokenOptions options = new TokenOptions();
                options.IamApiKey = key;
                string accessToken = new TokenManager(options).GetToken();
                if (string.IsNullOrEmpty(accessToken))
                {
                    throw new Exception(errorMessage);
                }
                if (!IsValidEndpoint(key, endpoint))
                {
                    throw new Exception(errorMessage);
                }
                Console.WriteLine("Token: {0}\n", accessToken);
            }
            catch
            {
                Console.WriteLine("Failed authentication.");
                return false;
            }
            return true;
        }
        public static void SaveStringToWaveFile(string text, string filePath, WatsonVoice voice)
        {
            if (!IsWatsonAccountPresentAndValid)
            {
                return;
            }
            string _key = WatsonAccount.GetInstance().GetKey();
            string _endpoint = EndpointToUriConverter.watsonRegionToEndpointMapping[WatsonAccount.GetInstance().GetRegion()];
            Text synthesizeText = new Text { _Text = text };
            TokenOptions options = new TokenOptions();
            options.IamApiKey = _key;
            string endpoint = _endpoint;
            var _service = new SynthesizeWatsonVoice(options, endpoint);

            var synthesizeResult = _service.Synthesize(synthesizeText, "audio/wav", voice: "en-US_" + voice.VoiceName);
            StreamUtility.SaveStreamToFile(filePath, synthesizeResult);
        }

        public static void Speak(string text, WatsonVoice voice)
        {
            string dirPath = Path.GetTempPath() + AudioService.TempFolderName;
            string filePath = dirPath + "\\" +
                string.Format(ELearningLabText.AudioPreviewFileNameFormat, voice.VoiceName);
            if (!Directory.Exists(dirPath))
            {
                Directory.CreateDirectory(dirPath);
            }
            SaveStringToWaveFile(text, filePath, voice);
            SpeechPlayingDialogBox speechPlayingDialog = new SpeechPlayingDialogBox();
            WaveOutEvent player = new WaveOutEvent();
            player.PlaybackStopped += (s, e) =>
            {
                try
                {
                    speechPlayingDialog.Dispatcher.Invoke(() => { speechPlayingDialog.Close(); });
                }
                catch
                {
                    Logger.Log("Object already disposed");
                }
            };
            speechPlayingDialog.Closed += (s, e) => SpeechPlayingDialog_Closed(player);
            try
            {
                using (var reader = new WaveFileReader(filePath))
                {
                    player.Init(reader);
                    player.Play();
                    speechPlayingDialog.ShowThematicDialog();
                }
            }
            catch
            {
                Logger.Log("Audio File not Found");
            }
        }

        private static ObservableCollection<WatsonVoice> GetWatsonVoices()
        {
            return new ObservableCollection<WatsonVoice>
            {
            new WatsonVoice(WatsonVoiceType.AllisonVoice),
            new WatsonVoice(WatsonVoiceType.LisaVoice),
            new WatsonVoice(WatsonVoiceType.MichaelVoice)
            };
        }

        private static void SpeechPlayingDialog_Closed(WaveOutEvent player)
        {
            player.Stop();
        }

        private static void SaveStringToWaveFile(string text, string filePath, WatsonVoice voice, string key, string endpoint)
        {
            Text synthesizeText = new Text { _Text = text };
            TokenOptions options = new TokenOptions();
            options.IamApiKey = key;
            var _service = new SynthesizeWatsonVoice(options, endpoint);

            var synthesizeResult = _service.Synthesize(synthesizeText, "audio/wav", voice: "en-US_" + voice.VoiceName);
            StreamUtility.SaveStreamToFile(filePath, synthesizeResult);
        }

        private static bool IsValidEndpoint(string key, string region)
        {
            string dirPath = Path.GetTempPath() + AudioService.TempFolderName;
            string filePath = dirPath + "\\" + ELearningLabText.WatsonAudioTestFileName;
            string textToSpeak = "This is to test watson voice.";
            if (!Directory.Exists(dirPath))
            {
                Directory.CreateDirectory(dirPath);
            }
            SaveStringToWaveFile(textToSpeak, filePath, new WatsonVoice(WatsonVoiceType.AllisonVoice), key, region);
            if (!File.Exists(filePath) || AudioService.ReadWavFileTimeSpan(filePath).TotalMilliseconds < 10)
            {
                Logger.Log("file exists? " + File.Exists(filePath).ToString());
                Logger.Log("Time span is " + AudioService.ReadWavFileTimeSpan(filePath).ToString());
                return false;
            }
            return true;
        }
    }
}
