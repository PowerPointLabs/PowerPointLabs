using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.Models;

namespace PowerPointLabs.ELearningLab.Service
{
    public class AzureRuntimeService
    {
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

        #region Audio Preview

        public static void SpeakString(string textToSpeak, AzureVoice voice)
        {
            string accessToken;
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
            azureVoiceSynthesizer.Speak(CancellationToken.None, new SynthesizeAzureVoice.InputOptions()
            {
                RequestUri = new Uri(requestUri),
                Text = textToSpeak,
                VoiceType = voice.voiceType,
                Locale = voice.Locale,
                VoiceName = voice.voiceName,
                // Service can return audio in different output format.
                OutputFormat = AudioOutputFormat.Riff24Khz16BitMonoPcm,
                AuthorizationToken = "Bearer " + accessToken,
            }).Wait();
        }

        #endregion

        #region Audio Generation
        public static bool SaveStringToWaveFileWithAzureVoice(string textToSave, string filePath, AzureVoice voice)
        {
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
            azureVoiceSynthesizer.Speak(CancellationToken.None, new SynthesizeAzureVoice.InputOptions()
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
            // For SoundPlayer to be able to play the wav file, it has to be encoded in PCM.
            // Use output audio format AudioOutputFormat.Riff16Khz16BitMonoPcm to do that.
            SoundPlayer player = new SoundPlayer(args.EventData);
            player.PlaySync();
            args.EventData.Dispose();
        }
    }
}
