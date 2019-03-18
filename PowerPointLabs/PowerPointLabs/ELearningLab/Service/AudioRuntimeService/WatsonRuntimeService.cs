using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NAudio.Wave;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Service
{
    public class WatsonRuntimeService
    {
        public static ObservableCollection<WatsonVoice> Voices { get; set; } = GetWatsonVoices();

        public static void SaveStringToWaveFile(string text, string filePath, WatsonVoice voice)
        {
            Text synthesizeText = new Text { _Text = text };
            TokenOptions options = new TokenOptions();

            var _service = new SynthesizeWatsonVoice(options, endpoint);
            Logger.Log(voice.VoiceName);

            var synthesizeResult = _service.Synthesize(synthesizeText, "audio/wav", voice: "en-US_" + voice.VoiceName);
            StreamUtility.SaveStreamToFile(filePath, synthesizeResult);
        }

        public static void Speak(string text, WatsonVoice voice)
        {
            string dirPath = System.IO.Path.GetTempPath() + AudioService.TempFolderName;
            string filePath = dirPath + "\\" +
                string.Format(ELearningLabText.AudioPreviewFileNameFormat, voice.VoiceName);
            ManualResetEvent syncEvent = new ManualResetEvent(false);
            Thread thread1 = new Thread(() =>
            {

                SaveStringToWaveFile(text, filePath, voice);
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
                };
                speechPlayingDialog.Closed += (s, e) => SpeechPlayingDialog_Closed(player);
                try
                {
                    using (var reader = new WaveFileReader(filePath))
                    {
                        player.Init(reader);
                        player.Play();
                        speechPlayingDialog.ShowDialog();
                    }
                }
                catch
                {
                    Logger.Log("Audio File not Found");
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
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
            Logger.Log("stoppppped.....");
            player.Stop();
        }
    }
}
