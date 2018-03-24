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
        private const long desiredFileSizeLimit = 75000L;

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

            Shape shapeToFillSlide = CompressImageInShape(pastingShape, targetCompression, desiredFileSizeLimit, tempPicPath, slide);

            shapeToFillSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            
            PPShape ppShapeToFillSlide = new PPShape(shapeToFillSlide);
            ppShapeToFillSlide.AbsoluteHeight = slideHeight;
            if (ppShapeToFillSlide.AbsoluteWidth < slideWidth)
            {
                ppShapeToFillSlide.AbsoluteWidth = slideWidth;
            }
            ppShapeToFillSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);
            
            CropLab.CropToSlide.Crop(shapeToFillSlide, slide, slideWidth, slideHeight);

            try
            {
                shapeToFillSlide.Select();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // do nothing, let user select
            }
        }

        private static Shape CompressImageInShape(Shape targetShape, long targetCompression, long fileSizeLimit, string tempFileStoragePath, PowerPointSlide currentSlide)
        {
            // Export the shape to extract the image within it
            targetShape.Export(tempFileStoragePath, PpShapeFormat.ppShapeFormatJPG);

            // Check if the image is acceptable in terms of size
            long fileLength = new FileInfo(tempFileStoragePath).Length;
            if (fileLength < fileSizeLimit)
            {
                // Delete the file as it is not needed anymore
                DeleteSpecificFilePath(tempFileStoragePath);

                // return the original shape
                return targetShape;
            }

            // Create a new bitmap from the image representing the exported shape
            Image img = Image.FromFile(tempFileStoragePath);
            Bitmap imgBitMap = new Bitmap(img);

            // Releases resources held by image object and delete temp file
            img.Dispose();
            DeleteSpecificFilePath(tempFileStoragePath);

            // Compresses and save the bitmap based on the specified level of compression (0 -> lowest quality, 100 -> highest quality)
            SaveJpg(imgBitMap, tempFileStoragePath, targetCompression);

            // Retrieve the compressed image and return a shape representing the image
            Shape compressedImgShape = currentSlide.Shapes.AddPicture(tempFileStoragePath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
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
    }
}
