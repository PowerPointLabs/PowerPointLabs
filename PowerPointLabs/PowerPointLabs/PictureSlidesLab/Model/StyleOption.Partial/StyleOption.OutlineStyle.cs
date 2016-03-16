using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseOutlineStyle { get; set; }

        [DefaultValue("#FFFFFF")]
        public string OutlineColor { get; set; }
    }
}
