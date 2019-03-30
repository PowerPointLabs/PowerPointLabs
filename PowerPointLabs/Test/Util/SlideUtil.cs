using System;
using System.IO;

using EyeOpen.Imaging.Processing;

using Microsoft.Office.Interop.PowerPoint;
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
                "different shape rotation for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.Rotation, actualShape.Rotation);
            Assert.IsTrue(IsRoughlySame(expShape.Left, actualShape.Left),
                "different shape left for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.Left, actualShape.Left);
            Assert.IsTrue(IsRoughlySame(expShape.Top, actualShape.Top),
                "different shape top for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.Top, actualShape.Top);
            Assert.IsTrue(IsRoughlySame(expShape.Width, actualShape.Width),
                "different shape width for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.Width, actualShape.Width);
            Assert.IsTrue(IsRoughlySame(expShape.Height, actualShape.Height),
                "different shape height for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.Height, actualShape.Height);
            Assert.AreEqual(expShape.HorizontalFlip, actualShape.HorizontalFlip,
                "different shape horizontal flip for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.HorizontalFlip, actualShape.HorizontalFlip);
            Assert.AreEqual(expShape.VerticalFlip, actualShape.VerticalFlip,
                "different shape vertical flip for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.VerticalFlip, actualShape.VerticalFlip);
        }

        public static void IsSameText(Shape expShape, Shape actualShape)
        {
            Assert.AreEqual(expShape.HasTextFrame, actualShape.HasTextFrame,
                "different text frame for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.HasTextFrame, actualShape.HasTextFrame);

            if (expShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                TextRange expText = expShape.TextFrame.TextRange;
                TextRange actualText = actualShape.TextFrame.TextRange;

                Assert.AreEqual(expText.Text, actualText.Text,
                    "different text for {0}. exp:{1}, actual:{2}", expShape.Name, expText.Text, actualText.Text);
                Assert.AreEqual(expText.Font.Color.RGB, actualText.Font.Color.RGB,
                    "different font color for {0}. exp:{1}, actual:{2}", expShape.Name, expText.Font.Color.RGB, actualText.Font.Color.RGB);
                Assert.AreEqual(expText.Font.Bold, actualText.Font.Bold,
                    "different bold style for {0}. exp:{1}, actual:{2}", expShape.Name, expText.Font.Bold, expText.Font.Bold);
                Assert.AreEqual(expText.Font.Underline, actualText.Font.Underline,
                    "different underline style for {0}. exp:{1}, actual:{2}", expShape.Name, expText.Font.Underline, expText.Font.Underline);
                Assert.AreEqual(expText.Font.Italic, actualText.Font.Italic,
                    "different italic style for {0}. exp:{1}, actual:{2}", expShape.Name, expText.Font.Italic, expText.Font.Italic);
                Assert.AreEqual(expText.Font.Size, actualText.Font.Size,
                    "different text size for {0}. exp:{1}, actual:{2}", expShape.Name, expText.Font.Size, expText.Font.Size);
            }
        }

        public static void IsSameZOrderPosition(Shape expShape, Shape actualShape)
        {
            Assert.AreEqual(expShape.ZOrderPosition, actualShape.ZOrderPosition,
                "different shape z order for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.ZOrderPosition, actualShape.ZOrderPosition);
        }

        public static void IsSameShapes(Slide expSlide, Slide actualSlide)
        {
            Shapes expShapes = expSlide.Shapes;
            Shapes actualShapes = actualSlide.Shapes;
            Assert.AreEqual(expShapes.Count, actualShapes.Count,
                "different number of shapes on slide. exp:{0}, actual:{1}", expShapes.Count, actualShapes.Count);

            for (int i = 1; i <= expShapes.Count; i++)
            {
                Shape expShape = expShapes[i];
                bool isMatchingShapeFound = false;
                for (int j = 1; j <= actualShapes.Count; j++)
                {
                    Shape actualShape = actualShapes[j];
                    if (expShape.Name == actualShape.Name)
                    {
                        isMatchingShapeFound = true;
                        IsSameShapeGroup(expShape, actualShape);
                        break;
                    }
                }
                Assert.IsTrue(isMatchingShapeFound, "no matching shape found");
            }
        }

        public static void IsSameShapeGroup(Shape expShape, Shape actualShape)
        {
            if (expShape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                Assert.AreEqual(expShape.Type, actualShape.Type,
                    "different shape type for {0}. exp:{1}, actual:{2}", expShape.Name, expShape.Type, actualShape.Type);
                Assert.AreEqual(expShape.GroupItems.Count, actualShape.GroupItems.Count,
                    "different number of shapes in group for {0}. exp:{1}, actual:{2}",
                    expShape.Name, expShape.GroupItems.Count, actualShape.GroupItems.Count);

                for (int i = 1; i <= expShape.GroupItems.Count; i++)
                {
                    IsSameShapeGroup(expShape.GroupItems[i], actualShape.GroupItems[i]);
                }
            }

            // even if shape type is group, we still want to check if the overall group is equal
            IsSameShape(expShape, actualShape);
            IsSameText(expShape, actualShape);
        }

        // compare shape's prop & looking
        //
        // You may need to call PpOperations.ExportSelectedShapes()
        // to get FileInfo of the exported shape in pic
        public static void IsSameLooking(Shape expShape, FileInfo expFileInfo, Shape actualShape,
            FileInfo actualFileInfo, double similarityTolerance = 0.95)
        {
            IsSameShape(expShape, actualShape);

            ComparableImage actualShapeInPic = new ComparableImage(actualFileInfo);
            ComparableImage expShapeInPic = new ComparableImage(expFileInfo);

            double similarity = actualShapeInPic.CalculateSimilarity(expShapeInPic);
            Assert.IsTrue(similarity > similarityTolerance, "The shapes look different. Similarity = " + similarity);
        }

        public static void IsSameLooking(Slide expSlide, Slide actualSlide, double similarityTolerance = 0.95)
        {
            int hashCode = DateTime.Now.GetHashCode();
            actualSlide.Export(PathUtil.GetTempPath("actualSlide" + hashCode + ".png"), "PNG");
            expSlide.Export(PathUtil.GetTempPath("expSlide" + hashCode + ".png"), "PNG");

            IsSameLooking("expSlide" + hashCode + ".png", "actualSlide" + hashCode + ".png", similarityTolerance);
        }

        public static void IsSameLooking(string expSlideImage, string actualSlideImage, double similarityTolerance = 0.95)
        {
            IsSameLooking(new FileInfo(PathUtil.GetTempPath(expSlideImage)), 
                new FileInfo(PathUtil.GetTempPath(actualSlideImage)),
                similarityTolerance);
        }

        public static void IsSameLooking(FileInfo expSlideImage, FileInfo actualSlideImage, double similarityTolerance = 0.95)
        {
            ComparableImage actualSlideInPic = new ComparableImage(actualSlideImage);
            ComparableImage expSlideInPic = new ComparableImage(expSlideImage);

            double similarity = actualSlideInPic.CalculateSimilarity(expSlideInPic);
            Assert.IsTrue(similarity > similarityTolerance, "The slides look different. Similarity = " + similarity);
        }

        public static void IsSameAnimations(Slide expSlide, Slide actualSlide)
        {
            Sequence actualSeq = actualSlide.TimeLine.MainSequence;
            Sequence expSeq = expSlide.TimeLine.MainSequence;
            Assert.AreEqual(expSeq.Count, actualSeq.Count, "Different animation sequence count.");
            for (int i = 1; i <= actualSeq.Count; i++)
            {
                Effect actualEffect = actualSeq[i];
                Effect expEffect = expSeq[i];
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

        public static bool IsAnimationsRemoved(Slide slide, string animPrefix)
        {
            Sequence slideSeq = slide.TimeLine.MainSequence;
            foreach (Effect effect in slideSeq)
            {
                if (effect.Shape.Name.Contains(animPrefix))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
