using System;
using System.Drawing;
using System.IO;
using System.Windows.Media.Imaging;

using ImageProcessor;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class ImageUtil
    {
        private const int ThumbnailHeight = 350;

        public static string GetThumbnailFromFullSizeImg(string filename)
        {
            if (filename == null)
            {
                return null;
            }

            string thumbnailPath = StoragePath.GetPath("thumbnail-"
                + DateTime.Now.GetHashCode() + "-"
                + Guid.NewGuid().ToString().Substring(0, 7));
            using (ImageFactory imageFactory = new ImageFactory())
            {
                Image image = imageFactory
                        .Load(filename)
                        .Image;
                float ratio = (float) image.Width / image.Height;
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
            using (ImageFactory imageFactory = new ImageFactory())
            {
                Image image = imageFactory
                        .Load(filename)
                        .Image;
                result = image.Width + " x " + image.Height;
            }
            return result;
        }
    }
}
