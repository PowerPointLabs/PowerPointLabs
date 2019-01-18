using System.Collections.Generic;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory.Options.Interface;

namespace PowerPointLabs.PictureSlidesLab.ModelFactory.Options
{
    abstract class BaseStyleOptions : IStyleOptions
    {
        public abstract List<StyleOption> GetOptionsForVariation();

        public abstract StyleOption GetDefaultOptionForPreview();

        protected List<StyleOption> GetOptions()
        {
            List<StyleOption> result = new List<StyleOption>();
            for (int i = 0; i < 8; i++)
            {
                result.Add(new StyleOption());
            }
            return result;
        }

        protected List<StyleOption> GetOptionsWithSuitableFontColor()
        {
            List<StyleOption> result = GetOptions();
            result[0].FontColor = "#000000"; //white(bg color) + black
            result[1].FontColor = "#FFD700"; //black + yellow
            result[2].FontColor = "#000000"; //yellow + black
            result[4].FontColor = "#001550"; //green + dark blue
            result[6].FontColor = "#FFD700"; //purple + yellow
            result[7].FontColor = "#3DFF8F"; //dark blue + green
            return result;
        }

        protected List<StyleOption> UpdateStyleName(List<StyleOption> opts, string styleName)
        {
            foreach (StyleOption styleOption in opts)
            {
                styleOption.StyleName = styleName;
            }
            return opts;
        }
    }
}
