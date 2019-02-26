using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Media;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using NAudio.Wave;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.Models;

namespace PowerPointLabs.ELearningLab.Service
{
    public class AzureRuntimeService
    {
        private static CancellationTokenSource cts = new CancellationTokenSource();
        private static CancellationToken token = cts.Token;
        public static bool IsAzureAccountPresent()
        {
            return !AzureAccount.GetInstance().IsEmpty();
        }

        public static bool IsValidUserAccount(bool showErrorMessage = true)
        {
            try
            {
                string _key = AzureAccount.GetInstance().GetKey();
                string _endpoint = AzureEndpointToUriConverter.regionToEndpointMapping[AzureAccount.GetInstance().GetRegion()];
                AzureAccountAuthentication auth = AzureAccountAuthentication.GetInstance(_endpoint, _key);
                string accessToken = auth.GetAccessToken();
                Console.WriteLine("Token: {0}\n", accessToken);
            }
            catch
            {
                Console.WriteLine("Failed authentication.");
                if (showErrorMessage)
                {
                    MessageBox.Show("Failed authentication");
                }
                return false;
            }
            return true;
        }

        public static bool IsValidUserAccount(string key, string endpoint)
        {
            try
            {
                AzureAccountAuthentication auth = AzureAccountAuthentication.GetInstance(endpoint, key);
                string accessToken = auth.GetAccessToken();
                Console.WriteLine("Token: {0}\n", accessToken);
            }
            catch
            {
                Console.WriteLine("Failed authentication.");
                return false;
            }
            return true;
        }

        public static void RenewCancellationToken()
        {
            cts = new CancellationTokenSource();
        }

        public static void Cancel()
        {
            cts.Cancel();
        }

        #region Audio Preview

        public static void SpeakString(string textToSpeak, AzureVoice voice)
        {
            RenewCancellationToken();
            string accessToken;
            string filePath = Path.GetTempPath() + AudioService.TempFolderName + "\\" + "PPTL_preview.wav";
            try
            {
                AzureAccountAuthentication auth = AzureAccountAuthentication.GetInstance();
                accessToken = auth.GetAccessToken();
                Logger.Log("Token: " + accessToken);
            }
            catch
            {
                Logger.Log("Failed authentication.");
                return;
            }
            string requestUri = AzureAccount.GetInstance().GetUri();
            if (requestUri == null)
            {
                return;
            }
            var azureVoiceSynthesizer = new SynthesizeAzureVoice();

            azureVoiceSynthesizer.OnAudioAvailable += PlayAudio;
            azureVoiceSynthesizer.OnError += OnAzureVoiceErrorHandler;
            // Reuse Synthesize object to minimize latency
            azureVoiceSynthesizer.Speak(token, new SynthesizeAzureVoice.InputOptions()
            {
                RequestUri = new Uri(requestUri),
                Text = textToSpeak,
                VoiceType = voice.voiceType,
                Locale = voice.Locale,
                VoiceName = voice.voiceName,
                // Service can return audio in different output format.
                OutputFormat = AudioOutputFormat.Riff24Khz16BitMonoPcm,
                AuthorizationToken = "Bearer " + accessToken,
            }, filePath).Wait();
        }

        #endregion

        #region Audio Generation
        public static bool SaveStringToWaveFileWithAzureVoice(string textToSave, string filePath, AzureVoice voice)
        {
            RenewCancellationToken();
            string accessToken;
            string textToSpeak = GetHumanSpeakNotesForText(textToSave);

            try
            {
                AzureAccountAuthentication auth = AzureAccountAuthentication.GetInstance();
                accessToken = auth.GetAccessToken();
                Logger.Log("Token: " + accessToken);
            }
            catch
            {
                Logger.Log("Failed authentication.");
                return false;
            }
            string requestUri = AzureAccount.GetInstance().GetUri();
            if (requestUri == null)
            {
                return false;
            }
            var azureVoiceSynthesizer = new SynthesizeAzureVoice();

            azureVoiceSynthesizer.OnAudioAvailable += SaveAudioToWaveFile;
            azureVoiceSynthesizer.OnError += OnAzureVoiceErrorHandler;

            // Reuse Synthesize object to minimize latency
            azureVoiceSynthesizer.Speak(token, new SynthesizeAzureVoice.InputOptions()
            {
                RequestUri = new Uri(requestUri),
                Text = textToSpeak,
                VoiceType = voice.voiceType,
                Locale = voice.Locale,
                VoiceName = voice.voiceName,
                // Service can return audio in different output format.
                OutputFormat = AudioOutputFormat.Riff24Khz16BitMonoPcm,
                AuthorizationToken = "Bearer " + accessToken,
            }, filePath).Wait();
            return true;
        }

        private static string GetHumanSpeakNotesForText(string textToSave)
        {
            TaggedText taggedText = new TaggedText(textToSave);
            string strToSpeak = taggedText.ToPrettyString();
            return strToSpeak;
        }

        private static void SaveAudioToWaveFile(object sender, GenericEventArgs<Stream> args)
        {
            SaveStreamToFile(args.FilePath, args.EventData);
            args.EventData.Dispose();
        }

        private static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            try
            {
                byte[] bytesInStream = ReadFully(stream);
                using (FileStream fs = File.Create(fileFullPath))
                {
                    fs.Write(bytesInStream, 0, bytesInStream.Length);
                }
            }
            catch
            {
                MessageBox.Show("Error generating audio files. ");
            }
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

        private static void OnAzureVoiceErrorHandler(object sender, GenericEventArgs<Exception> e)
        {
            Logger.Log("Unable to complete the TTS request: " + e.ToString());
        }
        #endregion

        private static void PlayAudio(object sender, GenericEventArgs<Stream> args)
        {
            ManualResetEvent syncEvent = new ManualResetEvent(false);
            Thread thread1 = new Thread(() =>
            {
                SaveAudioToWaveFile(sender, args);
                syncEvent.Set();
            });
            thread1.Start();
            Thread thread = new Thread(() =>
            {
                syncEvent.WaitOne();
                SpeechPlayingDialogBox speechPlayingDialog = new SpeechPlayingDialogBox();
                WaveOutEvent player = new WaveOutEvent();
                player.PlaybackStopped += (s, e) =>
                {
                    speechPlayingDialog.Dispatcher.Invoke(() => { speechPlayingDialog.Close(); });
                    args.EventData.Dispose();
                };
                speechPlayingDialog.Closed += (s, e) => SpeechPlayingDialog_Closed(player);
                using (var reader = new WaveFileReader(args.FilePath))
                {
                    player.Init(reader);
                    player.Play();
                    speechPlayingDialog.ShowDialog();
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private static SpeechPlayingDialogBox ShowSpeechCancelDialog(WaveOutEvent player)
        {
            SpeechPlayingDialogBox speechPlayingDialog = new SpeechPlayingDialogBox();
            speechPlayingDialog.Closed += (sender, e) => SpeechPlayingDialog_Closed(player);
            speechPlayingDialog.ShowDialog();
            return speechPlayingDialog;
        }

        private static void SpeechPlayingDialog_Closed(WaveOutEvent player)
        {
            player.Stop();
        }
    }
}
