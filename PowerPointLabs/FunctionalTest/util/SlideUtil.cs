using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using EyeOpen.Imaging.Processing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest.util
{
    class SlideUtil
    {
        // You may need to call PpOperations.ExportSelectedShapes()
        // to get FileInfo of the exported shape in pic
        public static void IsSameLooking(Shape expShape, FileInfo expFileInfo, Shape actualShape, FileInfo actualFileInfo)
        {
            Assert.AreEqual(expShape.Type, actualShape.Type);
            Assert.IsTrue(IsAlmostSame(expShape.Rotation, actualShape.Rotation), "different shape rotation");
            Assert.IsTrue(IsAlmostSame(expShape.Left, actualShape.Left), "different shape left");
            Assert.IsTrue(IsAlmostSame(expShape.Top, actualShape.Top), "different shape top");
            Assert.IsTrue(IsAlmostSame(expShape.Width, actualShape.Width), "different shape width");
            Assert.IsTrue(IsAlmostSame(expShape.Height, actualShape.Height), "different shape height");

            var actualShapeInPic = new ComparableImage(actualFileInfo);
            var expShapeInPic = new ComparableImage(expFileInfo);

            var similarity = actualShapeInPic.CalculateSimilarity(expShapeInPic);
            Assert.IsTrue(similarity > 0.99, "The shapes look different.");
        }

        public static void IsSameLooking(Slide expSlide, Slide actualSlide)
        {
            var hashCode = DateTime.Now.GetHashCode();
            actualSlide.Export(PathUtil.GetTempPath("actualSlide" + hashCode), "PNG");
            expSlide.Export(PathUtil.GetTempPath("expSlide" + hashCode), "PNG");

            var actualSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath("actualSlide" + hashCode)));
            var expSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath("expSlide" + hashCode)));

            var similarity = actualSlideInPic.CalculateSimilarity(expSlideInPic);
            Assert.IsTrue(similarity > 0.99, "The slides look different.");
        }

        public static void IsSameAnimations(Slide expSlide, Slide actualSlide)
        {
            var actualSeq = actualSlide.TimeLine.MainSequence;
            var expSeq = expSlide.TimeLine.MainSequence;
            Assert.AreEqual(expSeq.Count, actualSeq.Count, "Different animation sequence count.");
            for (int i = 1; i <= actualSeq.Count; i++)
            {
                var actualEffect = actualSeq[i];
                var expEffect = expSeq[i];
                // don't compare PPIndicator's effect
                Assert.AreEqual(expEffect.EffectType, actualEffect.EffectType, "Different effect type.");
                Assert.AreEqual(expEffect.Shape.Type, actualEffect.Shape.Type, "Different effect shape type.");
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Rotation, actualEffect.Shape.Rotation), "Different effect shape rotation.");
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Width, actualEffect.Shape.Width), "Different effect shape width.");
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Height, actualEffect.Shape.Height), "Different effect shape height.");
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Left, actualEffect.Shape.Left), "Different effect shape left.");
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Top, actualEffect.Shape.Top), "Different effect shape top.");
                Assert.IsTrue(IsAlmostSame(expEffect.Timing.TriggerDelayTime, actualEffect.Timing.TriggerDelayTime),
                    "Different effect timing.");
            }
        }

        private static bool IsAlmostSame(float a, float b)
        {
            return Math.Abs(a - b) < 0.001;
        }
    }
}
