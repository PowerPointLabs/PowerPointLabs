using System.Drawing;
using System.IO;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PasteLab
{
    static internal class PasteToFillSlide
    {
        private const long targetCompression = 95L;

        public static void Execute(PowerPointSlide slide, ShapeRange pastingShapes, float slideWidth, float slideHeight)
        {
            pastingShapes = ShapeUtil.GetShapesWhenTypeNotMatches(slide, pastingShapes, Microsoft.Office.Core.MsoShapeType.msoPlaceholder);
            if (pastingShapes.Count == 0)
            {
                return;
            }

            Shape pastingShape = pastingShapes[1];
            if (pastingShapes.Count > 1)
            {
                pastingShape = pastingShapes.Group();
            }

            string fileName = CommonText.TemporaryCompressedImageStorageFileName;
            string tempPicPath = Path.Combine(Path.GetTempPath(), fileName);

            Shape shapeToFillSlide = CompressImageInShape(pastingShape, targetCompression, tempPicPath, slide);

            shapeToFillSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            
            PPShape ppShapeToFillSlide = new PPShape(shapeToFillSlide);
            ppShapeToFillSlide.AbsoluteHeight = slideHeight;
            if (ppShapeToFillSlide.AbsoluteWidth < slideWidth)
            {
                ppShapeToFillSlide.AbsoluteWidth = slideWidth;
            }
            ppShapeToFillSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);
            
            CropLab.CropToSlide.Crop(shapeToFillSlide, slide, slideWidth, slideHeight);
            
            shapeToFillSlide.Select();
        }

        private static Shape CompressImageInShape(Shape targetShape, long targetCompression, string tempFileStoragePath, PowerPointSlide currentSlide)
        {
            // Create a new bitmap from the image represented by the imageShape
            targetShape.Export(tempFileStoragePath, PpShapeFormat.ppShapeFormatJPG);
            Image img = Image.FromFile(tempFileStoragePath);
            Bitmap imgBitMap = new Bitmap(img);

            // Releases resources held by image object and delete temp file
            img.Dispose();
            DeleteSpecificFilePath(tempFileStoragePath);

            // Compresses and save the bitmap based on the specified level of compression (0 -> lowest quality, 100 -> highest quality)
            SaveJpg(imgBitMap, tempFileStoragePath, targetCompression);

            // Retrieve the compressed image and return a shape representing the image
            Shape compressedImgShape = currentSlide.Shapes.AddPicture(tempFileStoragePath,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    targetShape.Left,
                    targetShape.Top);

            // Delete temp file again to return to original empty state
            DeleteSpecificFilePath(tempFileStoragePath);

            // Delete targetShape to prevent duplication
            targetShape.Delete();

            return compressedImgShape;
        }

        // Save the file with a specific compression level.
        private static void SaveJpg(Bitmap bm, string file_name, long compression)
        {
            try
            {
                System.Drawing.Imaging.EncoderParameters encoder_params = new System.Drawing.Imaging.EncoderParameters(1);
                encoder_params.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Quality, compression);

                System.Drawing.Imaging.ImageCodecInfo image_codec_info = GetEncoderInfo("image/jpeg");
                File.Delete(file_name);
                bm.Save(file_name, image_codec_info, encoder_params);
            }
            catch (System.Exception)
            {
            }
        }

        private static System.Drawing.Imaging.ImageCodecInfo GetEncoderInfo(string mime_type)
        {
            System.Drawing.Imaging.ImageCodecInfo[] encoders = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders();
            for (int i = 0; i <= encoders.Length; i++)
            {
                if (encoders[i].MimeType == mime_type)
                {
                    return encoders[i];
                }
            }
            return null;
        }

        private static void DeleteSpecificFilePath(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static System.Drawing.Image ScaleByPercent(System.Drawing.Image imgPhoto, int percent)
        {
            float nPercent = ((float)percent / 100);

            int sourceWidth = imgPhoto.Width;
            int sourceHeight = imgPhoto.Height;
            int sourceX = 0;
            int sourceY = 0;
            int destX = 0;
            int destY = 0;

            //Calcuate height and width of resized image.
            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = (int)(sourceHeight * nPercent);

            //Create a new bitmap object.
            System.Drawing.Bitmap bmPhoto = new System.Drawing.Bitmap(destWidth, destHeight,
                                     System.Drawing.Imaging.PixelFormat.Format24bppRgb);
            
            //Set resolution of bitmap.
            bmPhoto.SetResolution(imgPhoto.HorizontalResolution,
                                    imgPhoto.VerticalResolution);

            //Create a graphics object and set quality of graphics.
            System.Drawing.Graphics grPhoto = System.Drawing.Graphics.FromImage(bmPhoto);
            grPhoto.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            //Draw image by using DrawImage() method of graphics class.
            grPhoto.DrawImage(imgPhoto,
                new System.Drawing.Rectangle(destX, destY, destWidth, destHeight),
                new System.Drawing.Rectangle(sourceX, sourceY, sourceWidth, sourceHeight),
                System.Drawing.GraphicsUnit.Pixel);

            grPhoto.Dispose();   //Dispose graphics object.
            return bmPhoto;
        }

    }
}
