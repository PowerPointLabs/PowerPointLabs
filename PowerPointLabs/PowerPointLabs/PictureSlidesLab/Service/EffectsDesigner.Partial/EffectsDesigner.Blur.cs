using System;
using System.Drawing;

using ImageProcessor;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        private const float MinThumbnailHeight = 11f;
        private const float MaxThumbnailHeight = 1100f;

        public PowerPoint.Shape ApplyBlurEffect(string imageFileToBlur = null, int degree = 85)
        {
            Source.BlurImageFile = BlurImage(imageFileToBlur
                ?? Source.FullSizeImageFile
                ?? Source.ImageFile, degree);
            PowerPoint.Shape blurImageShape = AddPicture(Source.BlurImageFile, EffectName.Blur);
            float slideWidth = SlideWidth;
            float slideHeight = SlideHeight;
            FitToSlide.AutoFit(blurImageShape, slideWidth, slideHeight);
            CropPicture(blurImageShape);
            return blurImageShape;
        }

        private static string BlurImage(string imageFilePath, int degree)
        {
            if (degree == 0)
            {
                return imageFilePath;
            }

            string resizeImageFile = Util.TempPath.GetPath("fullsize_resize");
            using (ImageFactory imageFactory = new ImageFactory())
            {
                Image image = imageFactory
                    .Load(imageFilePath)
                    .Image;

                float ratio = (float)image.Width / image.Height;
                double targetHeight = Math.Round(MaxThumbnailHeight - (MaxThumbnailHeight - MinThumbnailHeight) / 100f * degree);
                double targetWidth = Math.Round(targetHeight * ratio);

                image = imageFactory
                    .Resize(new Size((int)targetWidth, (int)targetHeight))
                    .Image;
                image.Save(resizeImageFile);
            }

            string blurImageFile = Util.TempPath.GetPath("fullsize_blur");
            using (ImageFactory imageFactory = new ImageFactory())
            {
                Image image = imageFactory
                    .Load(resizeImageFile)
                    .GaussianBlur(5)
                    .Image;
                image.Save(blurImageFile);
            }
            return blurImageFile;
        }
    }
}
