using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class SelfExplanationClickItem : ClickItem, IEquatable<SelfExplanationClickItem>
    {
        #region public properties
        public bool IsCallout
        {
            get
            {
                return isCallout;
            }
            set
            {
                isCallout = (bool)value;
                NotifyPropertyChanged("IsCallout");
                NotifyPropertyChanged("IsDummyItem");
            }
        }
        public bool IsCaption
        {
            get
            {
                return isCaption;
            }
            set
            {
                isCaption = (bool)value;
                NotifyPropertyChanged("IsCaption");
                NotifyPropertyChanged("IsDummyItem");
            }
        }
        public bool IsVoice
        {
            get
            {
                return isVoice;
            }
            set
            {
                isVoice = (bool)value;
                NotifyPropertyChanged("IsVoice");
                NotifyPropertyChanged("IsDummyItem");
            }
        }

        public bool HasShortVersion
        {
            get
            {
                return hasShortVersion;
            }
            set
            {
                hasShortVersion = (bool)value;
                if (!hasShortVersion)
                {
                    return;
                }
                if (string.IsNullOrEmpty(calloutText.Trim()))
                {
                    calloutText = captionText;
                    NotifyPropertyChanged("CalloutText");
                }

            }
        }

        public string CalloutText
        {
            get
            {
                return calloutText;
            }
            set
            {
                calloutText = value;
                NotifyPropertyChanged("CalloutText");
            }
        }
        public string CaptionText
        {
            get
            {
                return captionText;
            }
            set
            {
                captionText = value;
                if (!hasShortVersion)
                {
                    CalloutText = value;
                }
                NotifyPropertyChanged("CaptionText");
            }
        }
        public string VoiceLabel
        {
            get
            {
                return voiceLabel;
            }
            set
            {
                voiceLabel = value;
                NotifyPropertyChanged("VoiceLabel");
            }
        }

        public int TriggerIndex
        {
            get
            {
                return (int)trigger;
            }
            set
            {
                trigger = (TriggerType)value;
                NotifyPropertyChanged("TriggerIndex");
            }
        }

        public bool IsDummyItem
        {
            get
            {
                return !IsCallout && !IsVoice && !IsCaption;
            }
        }

        public bool IsTriggerTypeComboBoxEnabled
        {
            get
            {
                return isTriggerTypeComboBoxEnabled;
            }
            set
            {
                isTriggerTypeComboBoxEnabled = (bool)value;
                NotifyPropertyChanged("IsTriggerTypeComboBoxEnabled");
            }
        }

        public bool IsEmpty
        {
            get
            {
                return string.IsNullOrEmpty(CaptionText.Trim())
                    && string.IsNullOrEmpty(CalloutText.Trim());
            }
        }

        public int TagNo
        {
            get
            {
                return tagNo;
            }
        }

        #endregion

        #region Attributes

        public int tagNo;

        private bool isCallout;
        private bool isCaption;
        private bool isVoice;
        private bool hasShortVersion;
        private bool isTriggerTypeComboBoxEnabled;

        public string calloutText;
        public string captionText;
        private string voiceLabel;

        private TriggerType trigger;

        #endregion

        public SelfExplanationClickItem(string captionText, string calloutText = "", string voiceLabel = "", bool isCallout = false,
            bool isCaption = false, bool isVoice = false, TriggerType trigger = TriggerType.WithPrevious,
            bool isTriggerTypeComboBoxEnabled = true, int tagNo = -1)
        {
            this.isCallout = isCallout;
            this.isCaption = isCaption;
            this.isVoice = isVoice;
            // we initailize callout text to be the same as caption text
            this.calloutText = string.IsNullOrEmpty(calloutText.Trim()) ? captionText : calloutText;
            this.captionText = captionText;
            this.voiceLabel = voiceLabel;
            this.trigger = trigger; // default to with previous
            this.isTriggerTypeComboBoxEnabled = isTriggerTypeComboBoxEnabled;
            this.tagNo = tagNo;
            hasShortVersion = !this.calloutText.Equals(this.captionText);
        }
        public override bool Equals(object other)
        {
            if (other == null || other.GetType() != GetType())
            {
                return false;
            }

            if (ReferenceEquals(other, this))
            {
                return true;
            }
            return Equals(other as SelfExplanationClickItem);
        }

        public bool Equals(SelfExplanationClickItem other)
        {
            return isCallout == other.isCallout
                && isCaption == other.isCaption
                && isVoice == other.isVoice
                && CalloutText.Equals(other.CalloutText)
                && CaptionText.Equals(other.CaptionText)
                && VoiceLabel.Equals(other.VoiceLabel)
                && ClickNo == other.ClickNo;
        }

        public override int GetHashCode()
        {
            var hashCode = -1571720738;
            hashCode = hashCode * -1521134295 + TriggerIndex.GetHashCode();
            hashCode = hashCode * -1521134295 + tagNo.GetHashCode();
            hashCode = hashCode * -1521134295 + isCallout.GetHashCode();
            hashCode = hashCode * -1521134295 + isCaption.GetHashCode();
            hashCode = hashCode * -1521134295 + isVoice.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(calloutText);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(captionText);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(voiceLabel);
            hashCode = hashCode * -1521134295 + ClickNo.GetHashCode();
            return hashCode;
        }
    }
}
