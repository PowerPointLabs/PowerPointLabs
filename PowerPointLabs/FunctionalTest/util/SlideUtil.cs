using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using EyeOpen.Imaging.Processing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest.util
{
    class SlideUtil
    {
        // only comparing shape's properties
        public static void IsSameShape(Shape expShape, Shape actualShape)
        {
            Assert.AreEqual(expShape.Type, actualShape.Type);
            Assert.IsTrue(IsRoughlySame(expShape.Rotation, actualShape.Rotation),
                "different shape rotation. exp:{0}, actual:{1}", expShape.Rotation, actualShape.Rotation);
            Assert.IsTrue(IsRoughlySame(expShape.Left, actualShape.Left),
                "different shape left. exp:{0}, actual:{1}", expShape.Left, actualShape.Left);
            Assert.IsTrue(IsRoughlySame(expShape.Top, actualShape.Top),
                "different shape top. exp:{0}, actual:{1}", expShape.Top, actualShape.Top);
            Assert.IsTrue(IsRoughlySame(expShape.Width, actualShape.Width),
                "different shape width. exp:{0}, actual:{1}", expShape.Width, actualShape.Width);
            Assert.IsTrue(IsRoughlySame(expShape.Height, actualShape.Height),
                "different shape height. exp:{0}, actual:{1}", expShape.Height, actualShape.Height);
        }

        // compare shape's prop & looking
        //
        // You may need to call PpOperations.ExportSelectedShapes()
        // to get FileInfo of the exported shape in pic
        public static void IsSameLooking(Shape expShape, FileInfo expFileInfo, Shape actualShape, FileInfo actualFileInfo)
        {
            IsSameShape(expShape, actualShape);

            var actualShapeInPic = new ComparableImage(actualFileInfo);
            var expShapeInPic = new ComparableImage(expFileInfo);

            var similarity = actualShapeInPic.CalculateSimilarity(expShapeInPic);
            Assert.IsTrue(similarity > 0.95, "The shapes look different. Similarity = " + similarity);
        }

        public static void IsSameLooking(Slide expSlide, Slide actualSlide)
        {
            var hashCode = DateTime.Now.GetHashCode();
            actualSlide.Export(PathUtil.GetTempPath("actualSlide" + hashCode), "PNG");
            expSlide.Export(PathUtil.GetTempPath("expSlide" + hashCode), "PNG");

            IsSameLooking("expSlide" + hashCode, "actualSlide" + hashCode);
        }

        public static void IsSameLooking(string expSlideImage, string actualSlideImage)
        {
            var actualSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath(actualSlideImage)));
            var expSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath(expSlideImage)));

            var similarity = actualSlideInPic.CalculateSimilarity(expSlideInPic);
            Assert.IsTrue(similarity > 0.95, "The slides look different. Similarity = " + similarity);
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
                Assert.AreEqual(expEffect.EffectType, actualEffect.EffectType, 
                    "Different effect type.");
                IsSameShape(expEffect.Shape, actualEffect.Shape);
                Assert.IsTrue(IsAlmostSame(expEffect.Timing.TriggerDelayTime, actualEffect.Timing.TriggerDelayTime),
                    "Different effect timing. exp:{0}, actual:{1}", expEffect.Timing.TriggerDelayTime, 
                    actualEffect.Timing.TriggerDelayTime);
            }
        }

        protected static bool IsSame(float a, float b, double threshold)
        {
            return Math.Abs(a - b) < threshold;
        }

        public static bool IsAlmostSame(float a, float b)
        {
            return IsSame(a, b, 0.005);
        }

        public static bool IsRoughlySame(float a, float b)
        {
            return IsSame(a, b, 0.5);
        }
    }
}
