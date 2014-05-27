using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.AudioMisc
{
    internal class Audio
    {
        public enum AudioType
        {
            Record,
            Auto
        }

        public const int GeneratedSamplingRate = 22050;
        public const int RecordedSamplingRate = 11025;
        public const int GeneratedBitRate = 16;
        public const int RecordedBitRate = 8;

        public string Name { get; set; }
        public int MatchSciptID { get; set; }
        public string SaveName { get; set; }
        public string Length { get; set; }
        public int LengthMillis { get; set; }
        public AudioType Type { get; set; }

        /// <summary>
        /// Default constructor.
        /// </summary>
        public Audio() {}

        public Audio(string name, string saveName, int matchScriptID)
        {
            Name = name;
            MatchSciptID = matchScriptID;
            SaveName = saveName;
            Length = AudioHelper.GetAudioLengthString(saveName);
            LengthMillis = AudioHelper.GetAudioLength(saveName);
            Type = AudioHelper.GetAudioType(saveName);
        }
    }
}
