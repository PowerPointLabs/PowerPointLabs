using System;
using System.Collections.Generic;

using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.TextCollection;
using TestInterface;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class ExplanationItem : ClickItem, IEquatable<ExplanationItem>, IExplanationItem
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsCallout);
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsDummyItem);
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsCaption);
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsDummyItem);
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsVoice);
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsDummyItem);
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
                    NotifyPropertyChanged(ELearningLabText.ExplanationItem_CalloutText);
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_CalloutText);
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_CaptionText);
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_VoiceLabel);
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsTriggerTypeComboBoxEnabled);
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsVoiceLabelInvalid);
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_TriggerIndex);
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
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsVoiceLabelInvalid);
                NotifyPropertyChanged(ELearningLabText.ExplanationItem_IsTriggerTypeComboBoxEnabled);
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

        public bool IsVoiceLabelInvalid
        {
            get
            {
                bool result = !AudioService.CheckIfVoiceExists(voiceLabel);
                return result;
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

        public ExplanationItem(string captionText, string calloutText = "", string voiceLabel = "", bool isCallout = false,
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
            return Equals(other as ExplanationItem);
        }

        public bool Equals(ExplanationItem other)
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
