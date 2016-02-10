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
            var blurImageShape = AddPicture(Source.BlurImageFile, EffectName.Blur);
            var slideWidth = SlideWidth;
            var slideHeight = SlideHeight;
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

            var blurImageFile = Util.TempPath.GetPath("fullsize_blur");
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                    .Load(imageFilePath)
                    .Image;
                var ratio = (float)image.Width / image.Height;
                var targetHeight = Math.Ceiling(MaxThumbnailHeight - (MaxThumbnailHeight - MinThumbnailHeight) / 100f * degree);
                var targetWidth = Math.Ceiling(targetHeight * ratio);

                image = imageFactory
                    .Resize(new Size((int)targetWidth, (int)targetHeight))
                    .GaussianBlur(5).Image;
                image.Save(blurImageFile);
            }
            return blurImageFile;
        }
    }
}
