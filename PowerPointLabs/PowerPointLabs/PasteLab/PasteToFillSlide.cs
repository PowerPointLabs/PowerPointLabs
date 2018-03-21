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
        private const float targetDPI = 96.0f;

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

            Shape shapeToFillSlide = null;

            string fileName = CommonText.TemporaryCompressedImageStorageFileName;
            string tempPicPath = Path.Combine(Path.GetTempPath(), fileName);
            
            pastingShape.Export(tempPicPath, PpShapeFormat.ppShapeFormatJPG);
            Image img = Image.FromFile(tempPicPath);
            Bitmap shapeBitMap = new Bitmap(img);

            img.Dispose();
            FileInfo file = new FileInfo(tempPicPath);
            if (file.Exists)
            {
                file.Delete();
            }
            
            // Add code to compress the slide here, using ShapeToBitmap method from GraphicsUtil.cs
            //System.Drawing.Bitmap shapeBitMap = GraphicsUtil.ShapeToBitmap(pastingShape);
            System.Diagnostics.Debug.WriteLine("Original resolution: " + shapeBitMap.HorizontalResolution);
            if (shapeBitMap.HorizontalResolution > targetDPI)
            {
                

                System.Diagnostics.Debug.WriteLine("Previous horizontal resolution: " + shapeBitMap.HorizontalResolution);
                //System.Diagnostics.Debug.WriteLine("Previous vertical resolution: " + shapeBitMap.VerticalResolution);
                //System.Diagnostics.Debug.WriteLine("Previous width: " + shapeBitMap.Size.Width);
                //System.Diagnostics.Debug.WriteLine("Previous height: " + shapeBitMap.Size.Height);
                shapeBitMap.SetResolution(targetDPI, targetDPI);
                System.Diagnostics.Debug.WriteLine("New horizontal resolution: " + shapeBitMap.HorizontalResolution);
                //System.Diagnostics.Debug.WriteLine("New vertical resolution: " + shapeBitMap.VerticalResolution);
                //System.Diagnostics.Debug.WriteLine("New width: " + shapeBitMap.Size.Width);
                //System.Diagnostics.Debug.WriteLine("New height: " + shapeBitMap.Size.Height);
                shapeBitMap.Save(tempPicPath);

                shapeToFillSlide = slide.Shapes.AddPicture(tempPicPath,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    pastingShape.Left,
                    pastingShape.Top);
                
                FileInfo file2 = new FileInfo(tempPicPath);
                if (file2.Exists)
                {
                    file2.Delete();
                }
                
                pastingShape.Delete();
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Accepted resolution: " + shapeBitMap.HorizontalResolution);
                shapeToFillSlide = pastingShape;
            }
            

            shapeToFillSlide.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            /*
            PPShape ppShapeToFillSlide = new PPShape(shapeToFillSlide);
            ppShapeToFillSlide.AbsoluteHeight = slideHeight;
            if (ppShapeToFillSlide.AbsoluteWidth < slideWidth)
            {
                ppShapeToFillSlide.AbsoluteWidth = slideWidth;
            }
            ppShapeToFillSlide.VisualCenter = new System.Drawing.PointF(slideWidth / 2, slideHeight / 2);
            
            CropLab.CropToSlide.Crop(shapeToFillSlide, slide, slideWidth, slideHeight);
            */
            //shapeToFillSlide.Select();
        }
        public static PPShape Resize(PPShape originalShape, float w, float h)
        {
            //Original Image attributes
            float originalWidth = originalShape.AbsoluteWidth;
            float originalHeight = originalShape.AbsoluteHeight;

            // Figure out the ratio
            double ratioX = (double)w / (double)originalWidth;
            double ratioY = (double)h / (double)originalHeight;
            // use whichever multiplier is smaller
            double ratio = ratioX < ratioY ? ratioX : ratioY;

            // now we can get the new height and width
            int newHeight = System.Convert.ToInt32(originalHeight * ratio);
            int newWidth = System.Convert.ToInt32(originalWidth * ratio);

            originalShape.AbsoluteWidth = newWidth;
            originalShape.AbsoluteHeight = newHeight;

            return originalShape;
            /*
            Image thumbnail = new System.Drawing.Bitmap(newWidth, newHeight);
            Graphics graphic = System.Drawing.Graphics.FromImage(thumbnail);

            graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphic.SmoothingMode = SmoothingMode.HighQuality;
            graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphic.CompositingQuality = CompositingQuality.HighQuality;

            graphic.Clear(Color.Transparent);
            graphic.DrawImage(originalImage, 0, 0, newWidth, newHeight);

            return thumbnail;
            */
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
