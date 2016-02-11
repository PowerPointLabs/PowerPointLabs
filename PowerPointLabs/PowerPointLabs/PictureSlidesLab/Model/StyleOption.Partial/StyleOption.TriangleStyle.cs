using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseTriangleStyle { get; set; }

        [DefaultValue("#000000")]
        public string TriangleColor { get; set; }

        [DefaultValue(0)]
        public int TriangleTransparency { get; set; }
    }
}
