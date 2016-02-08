using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseTextGlow { get; set; }

        [DefaultValue("")]
        public string TextGlowColor { get; set; }
    }
}
