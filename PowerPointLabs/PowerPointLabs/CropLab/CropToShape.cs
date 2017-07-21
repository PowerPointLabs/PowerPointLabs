using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

using PowerPointLabs.EffectsLab;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

using Office = Microsoft.Office.Core;
using Selection = Microsoft.Office.Interop.PowerPoint.Selection;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.CropLab
{
    public class CropToShape
    {
        private const string MessageBoxTitle = "Unable to crop";

        private static readonly string SlidePicture = Path.GetTempPath() + @"\slide.png";
        private static readonly string FillInBackgroundPicture = Path.GetTempPath() + @"\currentFillInBg.png";

        public static Shape Crop(PowerPointSlide currentSlide, Selection selection,
                                double magnifyRatio = 1.0, bool isInPlace = false, bool handleError = true)
        {
            var shapeRange = selection.ShapeRange;
            if (selection.HasChildShapeRange)
            {
                shapeRange = selection.ChildShapeRange;
            }

            var croppedShape = Crop(currentSlide, shapeRange, isInPlace: isInPlace, handleError: handleError);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }

            return croppedShape;
        }

        public static Shape Crop(PowerPointSlide currentSlide, ShapeRange shapeRange, 
                                double magnifyRatio = 1.0, bool isInPlace = false, bool handleError = true)
        {
            try
            {
                var hasManyShapes = shapeRange.Count > 1;
                var shape = hasManyShapes ? shapeRange.Group() : shapeRange[1];
                var left = shape.Left;
                var top = shape.Top;
                shape.Cut();
                shapeRange = currentSlide.Shapes.Paste();
                shapeRange.Left = left;
                shapeRange.Top = top;
                if (hasManyShapes)
                {
                    shapeRange = shapeRange.Ungroup();
                }

                TakeScreenshotProxy(currentSlide, shapeRange);

                var ungroupedRange = EffectsLabUtil.UngroupAllShapeRange(currentSlide, shapeRange);
                List<Shape> shapeList = new List<Shape>();

                for (int i = 1; i <= ungroupedRange.Count; i++)
                {
                    var filledShape = FillInShapeWithImage(currentSlide, SlidePicture, ungroupedRange[i], magnifyRatio, isInPlace);
                    shapeList.Add(filledShape);
                }
                
                var croppedRange = currentSlide.ToShapeRange(shapeList);
                var croppedShape = (croppedRange.Count == 1) ? croppedRange[1] : croppedRange.Group();

                return croppedShape;
            }
            catch (Exception e)
            {
                throw new CropLabException(e.Message);
            }
        }

        public static Shape FillInShapeWithImage(PowerPointSlide currentSlide, string imageFile, Shape shape, 
                                                double magnifyRatio = 1.0, bool isInPlace = false)
        {
            CreateFillInBackgroundForShape(imageFile, shape, magnifyRatio);
            shape.Fill.UserPicture(FillInBackgroundPicture);

            shape.Line.Visible = Office.MsoTriState.msoFalse;

            if (isInPlace)
            {
                return shape;
            }

            shape.Copy();
            var shapeToReturn = currentSlide.Shapes.Paste()[1];
            shape.Delete();
            return shapeToReturn;
        }

        public static Bitmap KiCut(Bitmap original, float startX, float startY, float width, float height,
                                    double magnifyRatio = 1.0)
        {
            if (original == null) { return null; }
            try
            {
                var outputImage = new Bitmap((int)width, (int)height, PixelFormat.Format32bppArgb);

                var inverseRatio = 1 / magnifyRatio;

                var newWidth = width * inverseRatio;
                var newHeight = height * inverseRatio;
                var newY = startY + (1 - inverseRatio) / 2 * width;
                var newX = startX + (1 - inverseRatio) / 2 * width;

                var inputGraphics = Graphics.FromImage(outputImage);
                inputGraphics.DrawImage(original,
                    new Rectangle(0, 0, (int)width, (int)height),
                    new Rectangle((int)newX, (int)newY, (int)newWidth, (int)newHeight),
                    GraphicsUnit.Pixel);
                inputGraphics.Dispose();

                return outputImage;
            }
            catch
            {
                return null;
            }
        }

        private static void CreateFillInBackgroundForShape(string imageFile, Shape shape, double magnifyRatio = 1.0)
        {
            using (var slideImage = (Bitmap)Image.FromFile(imageFile))
            {
                if (shape.Rotation == 0)
                {
                    CreateFillInBackground(shape, slideImage, magnifyRatio);
                }
                else
                {
                    CreateRotatedFillInBackground(shape, slideImage, magnifyRatio);
                }
            }
        }

        private static void CreateFillInBackground(Shape shape, Bitmap slideImage, double magnifyRatio = 1.0)
        {
            var croppedImage = KiCut(slideImage,
                shape.Left * GraphicsUtil.PictureExportingRatio,
                shape.Top * GraphicsUtil.PictureExportingRatio,
                shape.Width * GraphicsUtil.PictureExportingRatio,
                shape.Height * GraphicsUtil.PictureExportingRatio,
                magnifyRatio);
            croppedImage.Save(FillInBackgroundPicture, ImageFormat.Png);
        }

        private static void CreateRotatedFillInBackground(Shape shape, Bitmap slideImage, double magnifyRatio = 1.0)
        {
            var rotatedShape = new PPShape(shape, false);
            var topLeftPoint = new PointF(rotatedShape.ActualTopLeft.X * GraphicsUtil.PictureExportingRatio,
                rotatedShape.ActualTopLeft.Y * GraphicsUtil.PictureExportingRatio);

            Bitmap rotatedImage = new Bitmap(slideImage.Width, slideImage.Height);

            using (Graphics g = Graphics.FromImage(rotatedImage))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                using (System.Drawing.Drawing2D.Matrix mat = new System.Drawing.Drawing2D.Matrix())
                {
                    mat.Translate(-topLeftPoint.X, -topLeftPoint.Y);
                    mat.RotateAt(-shape.Rotation, topLeftPoint);

                    g.Transform = mat;
                    g.DrawImage(slideImage, new Rectangle(0, 0, slideImage.Width, slideImage.Height));
                }
            }

            var magnifiedImage = KiCut(rotatedImage, 0, 0, shape.Width * GraphicsUtil.PictureExportingRatio,
                shape.Height * GraphicsUtil.PictureExportingRatio, magnifyRatio);
            magnifiedImage.Save(FillInBackgroundPicture, ImageFormat.Png);
        }

        private static void TakeScreenshotProxy(PowerPointSlide currentSlide, ShapeRange shapeRange)
        {
            shapeRange.Visible = Office.MsoTriState.msoFalse;
            GraphicsUtil.ExportSlide(currentSlide, SlidePicture);
            shapeRange.Visible = Office.MsoTriState.msoTrue;
        }
    }
}
