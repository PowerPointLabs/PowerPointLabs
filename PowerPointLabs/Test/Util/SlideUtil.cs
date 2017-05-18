using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using EyeOpen.Imaging.Processing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Test.Util
{
    class SlideUtil
    {
        // only comparing shape's properties
        public static void IsSameShape(Shape expShape, Shape actualShape)
        {
            Assert.AreEqual(expShape.Type, actualShape.Type);
            if (expShape.Name.StartsWith("PowerPointLabs Speech"))
            {
                // Audio shape no need to compare size and position,
                // otherwise it causes bugs under different Dpi OS.
                return;
            }

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
            Assert.IsTrue(expShape.HorizontalFlip == actualShape.HorizontalFlip,
                "different shape horizontal flip. exp:{0}, actual:{1}", expShape.HorizontalFlip, actualShape.HorizontalFlip);
            Assert.IsTrue(expShape.VerticalFlip == actualShape.VerticalFlip,
                "different shape vertical flip. exp:{0}, actual:{1}", expShape.VerticalFlip, actualShape.VerticalFlip);
        }

        public static void IsSameText(Shape expShape, Shape actualShape)
        {
            if (expShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                var expText = expShape.TextFrame.TextRange;
                var actualText = actualShape.TextFrame.TextRange;

                Assert.IsTrue(expText.Text == actualText.Text,
                    "different text. exp:{0}, actual{1}", expText.Text, actualText.Text);
                Assert.IsTrue(expText.Font.Color.RGB == actualText.Font.Color.RGB,
                    "different font color. exp:{0}, actual:{1}", expText.Font.Color.RGB, actualText.Font.Color.RGB);
                Assert.IsTrue(expText.Font.Bold == actualText.Font.Bold,
                    "different bold style. exp:{0}, actual:{1}", expText.Font.Bold, expText.Font.Bold);
                Assert.IsTrue(expText.Font.Underline == actualText.Font.Underline,
                    "different underline style. exp:{0}, actual:{1}", expText.Font.Underline, expText.Font.Underline);
                Assert.IsTrue(expText.Font.Italic == actualText.Font.Italic,
                    "different italic style. exp:{0}, actual:{1}", expText.Font.Italic, expText.Font.Italic);
                Assert.IsTrue(expText.Font.Size == actualText.Font.Size,
                    "different text size. exp:{0}, actual:{1}", expText.Font.Size, expText.Font.Size);
            }
        }

        public static void IsSameZOrderPosition(Shape expShape, Shape actualShape)
        {
            Assert.IsTrue(expShape.ZOrderPosition == actualShape.ZOrderPosition,
                "different shape z order. exp:{0}, actual:{1}", expShape.ZOrderPosition, actualShape.ZOrderPosition);
        }

        public static void IsSameShapes(Slide expSlide, Slide actualSlide)
        {
            var expShapes = expSlide.Shapes;
            var actualShapes = actualSlide.Shapes;
            Assert.IsTrue(expShapes.Count == actualShapes.Count,
                "different number of shapes on slide. exp:{0}, actual:{1}", expShapes.Count, actualShapes.Count);

            for (int i = 1; i <= expShapes.Count; i++)
            {
                var expShape = expShapes[i];
                bool isMatchingShapeFound = false;
                for (int j = 1; j <= actualShapes.Count; j++)
                {
                    var actualShape = actualShapes[j];
                    if (expShape.Name == actualShape.Name)
                    {
                        isMatchingShapeFound = true;
                        if (expShapes[i].Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            Assert.IsTrue(expShape.Type == actualShape.Type,
                                "difNferent shape type. {2}. exp:{0}, actual:{1}", expShape.Type, actualShape.Type, expShape.Name);
                            Assert.IsTrue(expShape.GroupItems.Count == actualShape.GroupItems.Count,
                                "different number of shapes in group. exp:{0}, actual:{1}",
                                expShape.GroupItems.Count, actualShape.GroupItems.Count);

                            for (int k = 1; k <= expShape.GroupItems.Count; k++)
                            {
                                IsSameShape(expShape.GroupItems[k], actualShape.GroupItems[k]);
                                IsSameText(expShape.GroupItems[k], actualShape.GroupItems[k]);
                            }
                        }

                        // even if shape type is group, we still want to check if the overall group is equal
                        IsSameShape(expShape, actualShape);
                        break;
                    }
                }
                Assert.IsTrue(isMatchingShapeFound, "no matching shape found");
            }
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
            actualSlide.Export(PathUtil.GetTempPath("actualSlide" + hashCode + ".png"), "PNG");
            expSlide.Export(PathUtil.GetTempPath("expSlide" + hashCode + ".png"), "PNG");

            IsSameLooking("expSlide" + hashCode + ".png", "actualSlide" + hashCode + ".png");
        }

        public static void IsSameLooking(string expSlideImage, string actualSlideImage)
        {
            IsSameLooking(new FileInfo(PathUtil.GetTempPath(expSlideImage)), 
                new FileInfo(PathUtil.GetTempPath(actualSlideImage)));
        }

        public static void IsSameLooking(FileInfo expSlideImage, FileInfo actualSlideImage)
        {
            var actualSlideInPic = new ComparableImage(actualSlideImage);
            var expSlideInPic = new ComparableImage(expSlideImage);

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
