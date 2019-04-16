using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using TestInterface;

namespace Test.Util
{
    // Util class to help assert the data
    // from PpOperation.FetchPresentationData &
    // PpOperation.FetchCurrentPresentationData
    public class PresentationUtil
    {
        public static void AssertEqual(List<ISlideData> expectedSlides, List<ISlideData> actualSlides)
        {
            Assert.AreEqual(expectedSlides.Count, actualSlides.Count);
            for (int i = 0; i < expectedSlides.Count; ++i)
            {
                SlideDataUtil.AssertEqual(expectedSlides[i], actualSlides[i]);
            }
        }

        public class SlideDataUtil
        {
            public static void AssertEqual(ISlideData expected, ISlideData actual)
            {
                SlideUtil.IsSameLooking(expected.SlideImage, actual.SlideImage);
                Assert.AreEqual(expected.Animation.Count, actual.Animation.Count, "Different animation sequence count.");
                int count = expected.Animation.Count;
                for (int i = 0; i < count; ++i)
                {
                    EffectDataUtil.AssertEqual(expected.Animation[i], actual.Animation[i]);
                }
            }
        }

        public class EffectDataUtil
        {
            public static void AssertEqual(IEffectData expected, IEffectData actual)
            {
                Assert.AreEqual(expected.EffectType, actual.EffectType, "Different effect type.");
                Assert.AreEqual(expected.ShapeType, actual.ShapeType, "Different effect shape type.");
                Assert.IsTrue(SlideUtil.IsRoughlySame(expected.ShapeRotation, actual.ShapeRotation),
                    "Different effect shape rotation. exp:{0}, actual:{1}", expected.ShapeRotation, actual.ShapeRotation);
                Assert.IsTrue(SlideUtil.IsRoughlySame(expected.ShapeWidth, actual.ShapeWidth),
                    "Different effect shape width. exp:{0}, actual:{1}", expected.ShapeWidth, actual.ShapeWidth);
                Assert.IsTrue(SlideUtil.IsRoughlySame(expected.ShapeHeight, actual.ShapeHeight),
                    "Different effect shape height. exp:{0}, actual:{1}", expected.ShapeHeight, actual.ShapeHeight);
                Assert.IsTrue(SlideUtil.IsRoughlySame(expected.ShapeLeft, actual.ShapeLeft),
                    "Different effect shape left. exp:{0}, actual:{1}", expected.ShapeLeft, actual.ShapeLeft);
                Assert.IsTrue(SlideUtil.IsRoughlySame(expected.ShapeTop, actual.ShapeTop),
                    "Different effect shape top. exp:{0}, actual:{1}", expected.ShapeTop, actual.ShapeTop);
                Assert.IsTrue(SlideUtil.IsAlmostSame(expected.TimingTriggerDelayTime, actual.TimingTriggerDelayTime),
                    "Different effect timing. exp:{0}, actual:{1}", expected.TimingTriggerDelayTime, 
                    actual.TimingTriggerDelayTime);
            }
        }
    }
}
