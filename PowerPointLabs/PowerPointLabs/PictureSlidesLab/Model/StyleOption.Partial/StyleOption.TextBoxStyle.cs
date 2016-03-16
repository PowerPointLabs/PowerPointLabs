using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue(false)]
        public bool IsUseTextBoxStyle { get; set; }

        [DefaultValue("#000000")]
        public string TextBoxColor { get; set; }

        [DefaultValue(25)]
        public int TextBoxTransparency { get; set; }
    }
}
