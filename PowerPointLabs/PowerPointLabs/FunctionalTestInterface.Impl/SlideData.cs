using System;
using System.Collections.Generic;
using System.IO;
using FunctionalTestInterface;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

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
            slide.Export(Path.GetTempPath() + SlideImage, "PNG");

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
        public MsoAnimEffect EffectType { get; private set; }
        public MsoShapeType ShapeType { get; private set; }
        public float ShapeRotation { get; private set; }
        public float ShapeWidth { get; private set; }
        public float ShapeHeight { get; private set; }
        public float ShapeLeft { get; private set; }
        public float ShapeTop { get; private set; }
        public float TimingTriggerDelayTime { get; private set; }

        private EffectData(Effect effect)
        {
            EffectType = effect.EffectType;
            ShapeType = effect.Shape.Type;
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
