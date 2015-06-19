using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FunctionalTest.util;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest.models
{
    public class PresentationCompareData
    {
        public struct SlideData
        {
            public readonly string slideImage;
            public readonly List<EffectData> animationSequence;

            private SlideData(Slide slide)
            {
                slideImage = SlideUtil.SaveAsSlideImage(slide);

                animationSequence = new List<EffectData>();
                var seq = slide.TimeLine.MainSequence;
                for (int i = 1; i <= seq.Count; ++i)
                {
                    var effect = seq[i];
                    animationSequence.Add(new EffectData(effect));
                }
                return;
                animationSequence = slide.TimeLine.MainSequence
                                                  .Cast<Effect>()
                                                  .Select(effect => new EffectData(effect))
                                                  .ToList();
            }

            public static SlideData SaveSlideData(Slide slide)
            {
                return new SlideData(slide);
            }

            public static void AssertEqual(SlideData expected, SlideData actual)
            {
                SlideUtil.IsSameLooking(expected.slideImage, actual.slideImage);
                Assert.AreEqual(expected.animationSequence.Count, actual.animationSequence.Count, "Different animation sequence count.");
                int count = expected.animationSequence.Count;
                for (int i = 0; i < count; ++i)
                {
                    EffectData.AssertEqual(expected.animationSequence[i], actual.animationSequence[i]);
                }
            }
        }

        public struct EffectData
        {
            public readonly MsoAnimEffect EffectType;
            public readonly MsoShapeType ShapeType;
            public readonly float ShapeRotation;
            public readonly float ShapeWidth;
            public readonly float ShapeHeight;
            public readonly float ShapeLeft;
            public readonly float ShapeTop;
            public readonly float TimingTriggerDelayTime;

            public EffectData(Effect effect)
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

            public static void AssertEqual(EffectData expected, EffectData actual)
            {
                Assert.AreEqual(expected.EffectType, actual.EffectType, "Different effect type.");
                Assert.AreEqual(expected.ShapeType, actual.ShapeType, "Different effect shape type.");
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeRotation, actual.ShapeRotation), "Different effect shape rotation.");
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeWidth, actual.ShapeWidth), "Different effect shape width.");
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeHeight, actual.ShapeHeight), "Different effect shape height.");
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeLeft, actual.ShapeLeft), "Different effect shape left.");
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeTop, actual.ShapeTop), "Different effect shape top.");
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.TimingTriggerDelayTime, actual.TimingTriggerDelayTime),
                    "Different effect timing.");
            }

        }
    }
}
