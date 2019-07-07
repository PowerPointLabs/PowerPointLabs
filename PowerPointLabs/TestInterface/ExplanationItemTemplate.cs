using System;
using System.Collections.Generic;

namespace TestInterface
{
    [Serializable]
    public class ExplanationItemTemplate : IExplanationItem
    {
        public bool IsCallout { get; set; }
        public bool IsCaption { get; set; }
        public bool IsVoice { get; set; }
        public string VoiceLabel { get; set; }
        public bool HasShortVersion { get; set; }
        public string CaptionText { get; set; }

        public override bool Equals(object obj)
        {
            ExplanationItemTemplate other = obj as ExplanationItemTemplate;
            return other != null &&
                IsCallout == other.IsCallout &&
                IsCaption == other.IsCaption &&
                IsVoice == other.IsVoice &&
                VoiceLabel == other.VoiceLabel &&
                HasShortVersion == other.HasShortVersion; //&&
                //CaptionText == other.CaptionText;
        }

        public override int GetHashCode()
        {
            var hashCode = -788003819;
            hashCode = hashCode * -1521134295 + IsCallout.GetHashCode();
            hashCode = hashCode * -1521134295 + IsCaption.GetHashCode();
            hashCode = hashCode * -1521134295 + IsVoice.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(VoiceLabel);
            hashCode = hashCode * -1521134295 + HasShortVersion.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(CaptionText);
            return hashCode;
        }
    }
}
