using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using EyeOpen.Imaging.Processing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest.util
{
    class SlideComparer
    {
        // You may need to call PpOperations.ExportSelectedShapes()
        // to get FileInfo of the exported shape in pic
        public static void IsSameLooking(Shape expShape, FileInfo expFileInfo, Shape actualShape, FileInfo actualFileInfo)
        {
            Assert.AreEqual(expShape.Left, actualShape.Left);
            Assert.AreEqual(expShape.Top, actualShape.Top);
            Assert.AreEqual(expShape.Width, actualShape.Width);
            Assert.AreEqual(expShape.Height, actualShape.Height);

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
                if (!expEffect.DisplayName.StartsWith("PPIndicator"))
                {
                    Assert.AreEqual(expEffect.DisplayName, actualEffect.DisplayName, "Different effect display name.");
                }
                Assert.AreEqual(expEffect.EffectType, actualEffect.EffectType, "Different effect type.");
                if (!expEffect.Shape.Name.StartsWith("PPIndicator"))
                {
                    Assert.AreEqual(expEffect.Shape.Name, actualEffect.Shape.Name, "Different effect shape name.");
                }
                Assert.AreEqual(expEffect.Shape.Width, actualEffect.Shape.Width, "Different effect shape width.");
                Assert.AreEqual(expEffect.Shape.Height, actualEffect.Shape.Height, "Different effect shape height.");
                Assert.AreEqual(expEffect.Shape.Left, actualEffect.Shape.Left, "Different effect shape left.");
                Assert.AreEqual(expEffect.Shape.Top, actualEffect.Shape.Top, "Different effect shape top.");
                Assert.IsTrue(Math.Abs(expEffect.Timing.TriggerDelayTime - actualEffect.Timing.TriggerDelayTime) < 0.001,
                    "Different effect timing.");
            }
        }
    }
}
