using System.ComponentModel;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    public class Settings : WPF.Observable.Model
    {
        public Settings()
        {
            Init();
        }

        #region APIs

        public Aspect GetDefaultAspectWhenCustomize()
        {
            switch (DefaultAspectWhenCustomize)
            {
                case 0:
                    return Aspect.RecommendedAspect;
                case 1:
                default:
                    return Aspect.PictureAspect;
            }
        }

        public Alignment GetCitationTextBoxAlignment()
        {
            switch (CitationAlignment)
            {
                case 0:
                    return Alignment.Left;
                case 1:
                    return Alignment.Centre;
                case 2:
                default:
                    return Alignment.Right;
            }
        }

        #endregion

        private void Init()
        {
            foreach (PropertyDescriptor property in TypeDescriptor.GetProperties(this))
            {
                DefaultValueAttribute myAttribute = (DefaultValueAttribute) property
                    .Attributes[typeof(DefaultValueAttribute)];
                if (myAttribute != null)
                {
                    property.SetValue(this, myAttribute.Value);
                }
            }
        }

        #region UI related prop

        // Aspect

        private int _defaultAspectWhenCustomize;

        [DefaultValue(1)]
        public int DefaultAspectWhenCustomize
        {
            get
            {
                return _defaultAspectWhenCustomize;
            }
            set
            {
                _defaultAspectWhenCustomize = value;
                OnPropertyChanged("DefaultAspectWhenCustomize");
            }
        }

        // Citation

        private bool _isInsertCitation;

        [DefaultValue(false)]
        public bool IsInsertCitation
        {
            get
            {
                return _isInsertCitation;
            }
            set
            {
                _isInsertCitation = value;
                OnPropertyChanged("IsInsertCitation");
            }
        }

        private int _citationFontSize;

        [DefaultValue(10)]
        public int CitationFontSize
        {
            get
            {
                return _citationFontSize;
            }
            set
            {
                _citationFontSize = value;
                OnPropertyChanged("CitationFontSize");
            }
        }

        private string _citationFontColor;

        [DefaultValue("#FFFFFF")]
        public string CitationFontColor
        {
            get
            {
                return _citationFontColor;
            }
            set
            {
                _citationFontColor = value;
                OnPropertyChanged("CitationFontColor");
            }
        }

        private int _citationAlignment;

        [DefaultValue(0)]
        public int CitationAlignment
        {
            get
            {
                return _citationAlignment;
            }
            set
            {
                _citationAlignment = value;
                OnPropertyChanged("CitationAlignment");
            }
        }

        private bool _isUseCitationTextBox;

        [DefaultValue(false)]
        public bool IsUseCitationTextBox
        {
            get
            {
                return _isUseCitationTextBox;
            }
            set
            {
                _isUseCitationTextBox = value;
                OnPropertyChanged("IsUseCitationTextBox");
            }
        }

        private string _citationTextBoxColor;

        [DefaultValue("#000000")]
        public string CitationTextBoxColor
        {
            get
            {
                return _citationTextBoxColor;
            }
            set
            {
                _citationTextBoxColor = value;
                OnPropertyChanged("CitationTextBoxColor");
            }
        }

        private bool _isInsertCitationToNote;

        [DefaultValue(false)]
        public bool IsInsertCitationToNote
        {
            get
            {
                return _isInsertCitationToNote;
            }
            set
            {
                _isInsertCitationToNote = value;
                OnPropertyChanged("IsInsertCitationToNote");
            }
        }

        #endregion

        #region other info

        // any properties that not linked to UI should be put here

        #endregion
    }
}
