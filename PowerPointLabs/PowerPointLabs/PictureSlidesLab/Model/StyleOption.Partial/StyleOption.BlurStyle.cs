using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseBlurStyle { get; set; }

        [DefaultValue(0)]
        public int BlurDegree { get; set; }
    }
}
