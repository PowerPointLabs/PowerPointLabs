using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.AudioMisc
{
    public class Audio
    {
        public enum AudioType
        {
            Record,
            Auto
        }

        public string Name { get; set; }
        public int MatchSciptID { get; set; }
        public string SaveName { get; set; }
        public string Length { get; set; }
        public int LengthMillis { get; set; }
        public AudioType Type { get; set; }
    }
}
