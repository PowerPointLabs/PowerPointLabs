using System.ComponentModel;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        public Alignment GetCitationTextBoxAlignment()
        {
            switch (ImageReferenceAlignment)
            {
                case 0:
                    return Alignment.Auto;
                case 1:
                    return Alignment.Left;
                case 2:
                    return Alignment.Centre;
                // case 3:
                default:
                    return Alignment.Right;
            }
        }

        [DefaultValue(14)]
        public int CitationFontSize { get; set; }

        [DefaultValue(0)]
        public int ImageReferenceAlignment { get; set; }

        [DefaultValue("")]
        public string ImageReferenceTextBoxColor { get; set; }
    }
}
