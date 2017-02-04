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

    [Serializable]
    public class EffectData : IEffectData
    {
        public int EffectType { get; private set; }
        public int ShapeType { get; private set; }
        public float ShapeRotation { get; private set; }
        public float ShapeWidth { get; private set; }
        public float ShapeHeight { get; private set; }
        public float ShapeLeft { get; private set; }
        public float ShapeTop { get; private set; }
        public float TimingTriggerDelayTime { get; private set; }

        private EffectData(Effect effect)
        {
            EffectType = effect.EffectType.GetHashCode();
            ShapeType = effect.Shape.Type.GetHashCode();
            ShapeRotation = effect.Shape.Rotation;
            ShapeWidth = effect.Shape.Width;
            ShapeHeight = effect.Shape.Height;
            ShapeLeft = effect.Shape.Left;
            ShapeTop = effect.Shape.Top;
            TimingTriggerDelayTime = effect.Timing.TriggerDelayTime;
        }

        public static IEffectData FromEffect(Effect effect)
        {
            return new EffectData(effect);
        }
    }
}
