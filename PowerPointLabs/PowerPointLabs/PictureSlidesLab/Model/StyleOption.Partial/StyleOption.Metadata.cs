using System.ComponentModel;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        [DefaultValue("Default")]
        // used as tooltip for Variation Stage
        public string OptionName { get; set; }

        [DefaultValue("")]
        // used for Reload Styles
        public string StyleName { get; set; }
    }
}
