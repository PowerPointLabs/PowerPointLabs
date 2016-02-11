using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseFrameStyle { get; set; }

        [DefaultValue("#FFFFFF")]
        public string FrameColor { get; set; }

        [DefaultValue(30)]
        public int FrameTransparency { get; set; }
    }
}
