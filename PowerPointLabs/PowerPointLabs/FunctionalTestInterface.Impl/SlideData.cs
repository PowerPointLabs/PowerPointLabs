using System;
using System.Collections.Generic;

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
            int hashCode = DateTime.Now.GetHashCode();
            SlideImage = "slide" + hashCode;
            slide.Export(TempPath.GetTempTestFolder() + SlideImage, "PNG");

            Animation = new List<IEffectData>();
            Sequence seq = slide.TimeLine.MainSequence;
            for (int i = 1; i <= seq.Count; ++i)
            {
                Effect effect = seq[i];
                Animation.Add(EffectData.FromEffect(effect));
            }
        }

        public static ISlideData FromSlide(Slide slide)
        {
            return new SlideData(slide);
        }
    }
}
