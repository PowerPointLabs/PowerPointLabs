using System.Drawing;
using ImageProcessor;
using ImageProcessor.Imaging.Filters;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch.Handler.Effect
{
    class ImageProcessHelper
    {
        public static string BlurImage(string imageFilePath, bool isBlurForFullsize)
        {
            var blurImageFile = TempPath.GetPath("fullsize_blur");
            using (var imageFactory = new ImageFactory())
            {
                if (isBlurForFullsize)
                {// for full-size image, need to resize first
                    var image = imageFactory
                        .Load(imageFilePath)
                        .Image;
                    image = imageFactory
                        .Resize(new Size(image.Width / 4, image.Height / 4))
                        .GaussianBlur(5).Image;
                    image.Save(blurImageFile);
                }
                else
                {
                    var image = imageFactory
                        .Load(imageFilePath)
                        .GaussianBlur(5)
                        .Image;
                    image.Save(blurImageFile);
                }
            }
            return blurImageFile;
        }

        public static string GrayscaleImage(string imageFilePath)
        {
            var grayscaleImageFile = TempPath.GetPath("fullsize_grayscale");
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                    .Load(imageFilePath)
                    .Filter(MatrixFilters.GreyScale)
                    .Image;
                image.Save(grayscaleImageFile);
            }
            return grayscaleImageFile;
        }
    }
}
