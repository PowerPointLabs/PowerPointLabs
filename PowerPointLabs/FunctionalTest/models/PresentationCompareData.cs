using System.Collections.Generic;
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
            private readonly string _slideImage;
            private readonly List<EffectData> _animationSequence;

            private SlideData(Slide slide)
            {
                _slideImage = SlideUtil.SaveAsSlideImage(slide);

                _animationSequence = new List<EffectData>();
                var seq = slide.TimeLine.MainSequence;
                for (int i = 1; i <= seq.Count; ++i)
                {
                    var effect = seq[i];
                    _animationSequence.Add(new EffectData(effect));
                }
            }

            public static SlideData SaveSlideData(Slide slide)
            {
                return new SlideData(slide);
            }

            public static void AssertEqual(SlideData expected, SlideData actual)
            {
                SlideUtil.IsSameLooking(expected._slideImage, actual._slideImage);
                Assert.AreEqual(expected._animationSequence.Count, actual._animationSequence.Count, "Different animation sequence count.");
                int count = expected._animationSequence.Count;
                for (int i = 0; i < count; ++i)
                {
                    EffectData.AssertEqual(expected._animationSequence[i], actual._animationSequence[i]);
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
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeRotation, actual.ShapeRotation),
                    "Different effect shape rotation. exp:{0}, actual:{1}", expected.ShapeRotation, actual.ShapeRotation);
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeWidth, actual.ShapeWidth),
                    "Different effect shape width. exp:{0}, actual:{1}", expected.ShapeWidth, actual.ShapeWidth);
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeHeight, actual.ShapeHeight),
                    "Different effect shape height. exp:{0}, actual:{1}", expected.ShapeHeight, actual.ShapeHeight);
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeLeft, actual.ShapeLeft),
                    "Different effect shape left. exp:{0}, actual:{1}", expected.ShapeLeft, actual.ShapeLeft);
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.ShapeTop, actual.ShapeTop),
                    "Different effect shape top. exp:{0}, actual:{1}", expected.ShapeTop, actual.ShapeTop);
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.TimingTriggerDelayTime, actual.TimingTriggerDelayTime),
                    "Different effect timing. exp:{0}, actual:{1}", expected.TimingTriggerDelayTime, 
                    actual.TimingTriggerDelayTime);
            }

        }
    }
}
