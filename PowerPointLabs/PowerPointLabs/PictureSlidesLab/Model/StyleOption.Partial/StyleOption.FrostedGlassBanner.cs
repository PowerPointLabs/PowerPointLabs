using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseFrostedGlassBannerStyle { get; set; }

        [DefaultValue("#000000")]
        public string FrostedGlassBannerColor { get; set; }

        [DefaultValue(80)]
        public int FrostedGlassBannerTransparency { get; set; }
    }
}
