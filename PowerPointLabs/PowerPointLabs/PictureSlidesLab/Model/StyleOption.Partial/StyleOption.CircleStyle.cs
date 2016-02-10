using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseCircleStyle { get; set; }

        [DefaultValue("#FFFFFF")]
        public string CircleColor { get; set; }

        [DefaultValue(0)]
        public int CircleTransparency { get; set; }
    }
}
