using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class InputOptions
    {
        public InputOptions()
        {
            this.Locale = "en-us";
            this.VoiceName = "Microsoft Server Speech Text to Speech Voice (en-US, ZiraRUS)";
            // Default to Riff16Khz16BitMonoPcm output format.
            this.OutputFormat = AudioOutputFormat.Riff16Khz16BitMonoPcm;
        }

        public Uri RequestUri { get; set; }
        public AudioOutputFormat OutputFormat { get; set; }

        public IEnumerable<KeyValuePair<string, string>> Headers
        {
            get
            {
                List<KeyValuePair<string, string>> toReturn = new List<KeyValuePair<string, string>>();
                toReturn.Add(new KeyValuePair<string, string>("Content-Type", "application/ssml+xml"));

                string outputFormat;

                switch (this.OutputFormat)
                {
                    case AudioOutputFormat.Raw16Khz16BitMonoPcm:
                        outputFormat = "raw-16khz-16bit-mono-pcm";
                        break;

                    case AudioOutputFormat.Raw8Khz8BitMonoMULaw:
                        outputFormat = "raw-8khz-8bit-mono-mulaw";
                        break;

                    case AudioOutputFormat.Riff16Khz16BitMonoPcm:
                        outputFormat = "riff-16khz-16bit-mono-pcm";
                        break;

                    case AudioOutputFormat.Riff8Khz8BitMonoMULaw:
                        outputFormat = "riff-8khz-8bit-mono-mulaw";
                        break;

                    case AudioOutputFormat.Ssml16Khz16BitMonoSilk:
                        outputFormat = "ssml-16khz-16bit-mono-silk";
                        break;

                    case AudioOutputFormat.Raw16Khz16BitMonoTrueSilk:
                        outputFormat = "raw-16khz-16bit-mono-truesilk";
                        break;

                    case AudioOutputFormat.Ssml16Khz16BitMonoTts:
                        outputFormat = "ssml-16khz-16bit-mono-tts";
                        break;

                    case AudioOutputFormat.Audio16Khz128KBitRateMonoMp3:
                        outputFormat = "audio-16khz-128kbitrate-mono-mp3";
                        break;

                    case AudioOutputFormat.Audio16Khz64KBitRateMonoMp3:
                        outputFormat = "audio-16khz-64kbitrate-mono-mp3";
                        break;

                    case AudioOutputFormat.Audio16Khz32KBitRateMonoMp3:
                        outputFormat = "audio-16khz-32kbitrate-mono-mp3";
                        break;

                    case AudioOutputFormat.Audio16Khz16KbpsMonoSiren:
                        outputFormat = "audio-16khz-16kbps-mono-siren";
                        break;

                    case AudioOutputFormat.Riff16Khz16KbpsMonoSiren:
                        outputFormat = "riff-16khz-16kbps-mono-siren";
                        break;
                    case AudioOutputFormat.Raw24Khz16BitMonoPcm:
                        outputFormat = "raw-24khz-16bit-mono-pcm";
                        break;
                    case AudioOutputFormat.Riff24Khz16BitMonoPcm:
                        outputFormat = "riff-24khz-16bit-mono-pcm";
                        break;
                    case AudioOutputFormat.Audio24Khz48KBitRateMonoMp3:
                        outputFormat = "audio-24khz-48kbitrate-mono-mp3";
                        break;
                    case AudioOutputFormat.Audio24Khz96KBitRateMonoMp3:
                        outputFormat = "audio-24khz-96kbitrate-mono-mp3";
                        break;
                    case AudioOutputFormat.Audio24Khz160KBitRateMonoMp3:
                        outputFormat = "audio-24khz-160kbitrate-mono-mp3";
                        break;
                    default:
                        outputFormat = "riff-16khz-16bit-mono-pcm";
                        break;
                }

                toReturn.Add(new KeyValuePair<string, string>("X-Microsoft-OutputFormat", outputFormat));
                // authorization Header
                toReturn.Add(new KeyValuePair<string, string>("Authorization", this.AuthorizationToken));
                // Refer to the doc
                toReturn.Add(new KeyValuePair<string, string>("X-Search-AppId", "07D3234E49CE426DAA29772419F436CA"));
                // Refer to the doc
                toReturn.Add(new KeyValuePair<string, string>("X-Search-ClientID", "1ECFAE91408841A480F00935DC390960"));
                // The software originating the request
                toReturn.Add(new KeyValuePair<string, string>("User-Agent", "TTSClient"));

                return toReturn;
            }
            set
            {
                Headers = value;
            }
        }

        public string Locale { get; set; }
        public Gender VoiceType { get; set; }
        public string VoiceName { get; set; }
        public string AuthorizationToken { get; set; }
        public string Text { get; set; }
    }
}
