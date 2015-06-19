using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using EyeOpen.Imaging.Processing;
using FunctionalTest.models;
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
            Assert.IsTrue(IsAlmostSame(expShape.Rotation, actualShape.Rotation), 
                "different shape rotation. exp:{0}, actual:{1}", expShape.Rotation, actualShape.Rotation);
            Assert.IsTrue(IsAlmostSame(expShape.Left, actualShape.Left), 
                "different shape left. exp:{0}, actual:{1}", expShape.Left, actualShape.Left);
            Assert.IsTrue(IsAlmostSame(expShape.Top, actualShape.Top),
                "different shape top. exp:{0}, actual:{1}", expShape.Top, actualShape.Top);
            Assert.IsTrue(IsAlmostSame(expShape.Width, actualShape.Width),
                "different shape width. exp:{0}, actual:{1}", expShape.Width, actualShape.Width);
            Assert.IsTrue(IsAlmostSame(expShape.Height, actualShape.Height),
                "different shape height. exp:{0}, actual:{1}", expShape.Height, actualShape.Height);

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

            var actualSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath("actualSlide" + hashCode)));
            var expSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath("expSlide" + hashCode)));

            var similarity = actualSlideInPic.CalculateSimilarity(expSlideInPic);
            Assert.IsTrue(similarity > 0.95, "The slides look different. Similarity = " + similarity);
        }

        public static void IsSameLooking(string expSlideImage, string actualSlideImage)
        {
            var actualSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath(actualSlideImage)));
            var expSlideInPic = new ComparableImage(new FileInfo(PathUtil.GetTempPath(expSlideImage)));

            var similarity = actualSlideInPic.CalculateSimilarity(expSlideInPic);
            Assert.IsTrue(similarity > 0.95, "The slides look different. Similarity = " + similarity);
        }

        public static string SaveAsSlideImage(Slide slide)
        {
            var hashCode = DateTime.Now.GetHashCode();
            string fileName = "slide" + hashCode;
            slide.Export(PathUtil.GetTempPath(fileName), "PNG");
            return fileName;
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
                Assert.AreEqual(expEffect.EffectType, actualEffect.EffectType, 
                    "Different effect type.");
                Assert.AreEqual(expEffect.Shape.Type, actualEffect.Shape.Type,
                    "Different effect shape type.");
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Rotation, actualEffect.Shape.Rotation),
                    "Different effect shape rotation. exp:{0}, actual:{1}", expEffect.Shape.Rotation, actualEffect.Shape.Rotation);
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Width, actualEffect.Shape.Width),
                    "Different effect shape width. exp:{0}, actual:{1}", expEffect.Shape.Width, actualEffect.Shape.Width);
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Height, actualEffect.Shape.Height),
                    "Different effect shape height. exp:{0}, actual:{1}", expEffect.Shape.Height, actualEffect.Shape.Height);
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Left, actualEffect.Shape.Left),
                    "Different effect shape left. exp:{0}, actual:{1}", expEffect.Shape.Left, actualEffect.Shape.Left);
                Assert.IsTrue(IsAlmostSame(expEffect.Shape.Top, actualEffect.Shape.Top),
                    "Different effect shape top. exp:{0}, actual:{1}", expEffect.Shape.Top, actualEffect.Shape.Top);
                Assert.IsTrue(IsAlmostSame(expEffect.Timing.TriggerDelayTime, actualEffect.Timing.TriggerDelayTime),
                    "Different effect timing. exp:{0}, actual:{1}", expEffect.Timing.TriggerDelayTime, 
                    actualEffect.Timing.TriggerDelayTime);
            }
        }

        public static bool IsAlmostSame(float a, float b)
        {
            return Math.Abs(a - b) < 0.005;
        }
    }
}
