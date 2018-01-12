using System;
using System.Collections.Generic;
using System.IO;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;
using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    public class SlideData : ISlideData
    {
        // path to slide's image
        public string SlideImage { get; private set; }
        public List<IEffectData> Animation { get; private set; }

        private SlideData(Slide slide)
        {
            var hashCode = DateTime.Now.GetHashCode();
            SlideImage = "slide" + hashCode;
            slide.Export(TempPath.GetTempTestFolder() + SlideImage, "PNG");

            Animation = new List<IEffectData>();
            var seq = slide.TimeLine.MainSequence;
            for (int i = 1; i <= seq.Count; ++i)
            {
                var effect = seq[i];
                Animation.Add(EffectData.FromEffect(effect));
            }
        }

        public static ISlideData FromSlide(Slide slide)
        {
            return new SlideData(slide);
        }
    }
}
