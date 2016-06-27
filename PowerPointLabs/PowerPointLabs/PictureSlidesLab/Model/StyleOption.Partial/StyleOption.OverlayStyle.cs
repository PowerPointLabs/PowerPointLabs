using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseOverlayStyle { get; set; }

        [DefaultValue("#000000")]
        public string OverlayColor { get; set; }

        // for background's overlay
        [DefaultValue(100)]
        public int OverlayTransparency { get; set; }
    }
}
