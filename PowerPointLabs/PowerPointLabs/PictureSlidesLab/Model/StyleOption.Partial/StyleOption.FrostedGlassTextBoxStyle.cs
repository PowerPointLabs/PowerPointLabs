using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseFrostedGlassTextBoxStyle { get; set; }

        [DefaultValue("#000000")]
        public string FrostedGlassTextBoxColor { get; set; }

        [DefaultValue(80)]
        public int FrostedGlassTextBoxTransparency { get; set; }
    }
}
