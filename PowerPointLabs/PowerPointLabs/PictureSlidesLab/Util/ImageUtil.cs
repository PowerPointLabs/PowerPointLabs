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

        public static BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }
    }
}
