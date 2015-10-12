using System;
using System.Drawing;
using ImageProcessor;

namespace PowerPointLabs.ImageSearch.Util
{
    class ImageUtil
    {
        private const int ThumbnailHeight = 350;

        public static string GetThumbnailFromFullSizeImg(string filename)
        {
            var thumbnailPath = StoragePath.GetPath("thumbnail-"
                + DateTime.Now.GetHashCode() + "-"
                + Guid.NewGuid().ToString().Substring(0, 7));
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                        .Load(filename)
                        .Image;
                var ratio = (float) image.Width / image.Height;
                image = imageFactory
                        .Resize(new Size((int)(ThumbnailHeight * ratio), ThumbnailHeight))
                        .Image;
                image.Save(thumbnailPath);
            }
            return thumbnailPath;
        }

        public static string GetWidthAndHeight(string filename)
        {
            string result;
            using (var imageFactory = new ImageFactory())
            {
                var image = imageFactory
                        .Load(filename)
                        .Image;
                result = image.Width + " x " + image.Height;
            }
            return result;
        }
    }
}
