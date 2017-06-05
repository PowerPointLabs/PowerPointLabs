using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace PowerPointLabs.CropLab
{
    public class CropToShape
    {
        private const string MessageBoxTitle = "Unable to crop";

        private static readonly string SlidePicture = Path.GetTempPath() + @"\slide.png";
        private static readonly string FillInBackgroundPicture = Path.GetTempPath() + @"\currentFillInBg.png";

        public static PowerPoint.Shape Crop(PowerPointSlide currentSlide, PowerPoint.Selection selection, double magnifyRatio = 1.0, bool isInPlace = false,
                                            bool handleError = true)
        {
            var croppedShape = Crop(currentSlide, selection.ShapeRange, isInPlace: isInPlace, handleError: handleError);
            if (croppedShape != null)
            {
                croppedShape.Select();
            }

            return croppedShape;
        }

        public static PowerPoint.Shape Crop(PowerPointSlide currentSlide, PowerPoint.ShapeRange shapeRange, double magnifyRatio = 1.0, bool isInPlace = false,
            bool handleError = true)
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

                var ungroupedRange = UngroupAllForShapeRange(currentSlide, shapeRange);
                var shapeNames = new string[ungroupedRange.Count];

                for (int i = 1; i <= ungroupedRange.Count; i++)
                {
                    var filledShape = FillInShapeWithImage(currentSlide, SlidePicture, ungroupedRange[i], magnifyRatio, isInPlace);
                    shapeNames[i - 1] = filledShape.Name;
                }

                var croppedRange = currentSlide.Shapes.Range(shapeNames);
                var croppedShape = (croppedRange.Count == 1) ? croppedRange[1] : croppedRange.Group();

                return croppedShape;
            }
            catch (Exception)
            {
                throw new CropLabException(TextCollection.CropToShapeText.ErrorMessageForUndefined);
            }
        }

        public static PowerPoint.Shape FillInShapeWithImage(PowerPointSlide currentSlide, string imageFile, PowerPoint.Shape shape, double magnifyRatio = 1.0,
            bool isInPlace = false)
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

        private static void CreateFillInBackgroundForShape(string imageFile, PowerPoint.Shape shape, double magnifyRatio = 1.0)
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

        private static void CreateFillInBackground(PowerPoint.Shape shape, Bitmap slideImage, double magnifyRatio = 1.0)
        {
            var croppedImage = KiCut(slideImage,
                shape.Left * Utils.Graphics.PictureExportingRatio,
                shape.Top * Utils.Graphics.PictureExportingRatio,
                shape.Width * Utils.Graphics.PictureExportingRatio,
                shape.Height * Utils.Graphics.PictureExportingRatio,
                magnifyRatio);
            croppedImage.Save(FillInBackgroundPicture, ImageFormat.Png);
        }

        private static void CreateRotatedFillInBackground(PowerPoint.Shape shape, Bitmap slideImage, double magnifyRatio = 1.0)
        {
            var rotatedShape = new Utils.PPShape(shape, false);
            var topLeftPoint = new PointF(rotatedShape.ActualTopLeft.X * Utils.Graphics.PictureExportingRatio,
                rotatedShape.ActualTopLeft.Y * Utils.Graphics.PictureExportingRatio);

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

            var magnifiedImage = KiCut(rotatedImage, 0, 0, shape.Width * Utils.Graphics.PictureExportingRatio,
                shape.Height * Utils.Graphics.PictureExportingRatio, magnifyRatio);
            magnifiedImage.Save(FillInBackgroundPicture, ImageFormat.Png);
        }

        private static void TakeScreenshotProxy(PowerPointSlide currentSlide, PowerPoint.ShapeRange shapeRange)
        {
            shapeRange.Visible = Office.MsoTriState.msoFalse;
            Utils.Graphics.ExportSlide(currentSlide, SlidePicture);
            shapeRange.Visible = Office.MsoTriState.msoTrue;
        }

        private static PowerPoint.ShapeRange UngroupAllForShapeRange(PowerPointSlide currentSlide, PowerPoint.ShapeRange range)
        {
            var ungroupedShapeNames = new List<string>();
            var queue = new Queue<PowerPoint.Shape>();

            foreach (var item in range)
            {
                queue.Enqueue(item as PowerPoint.Shape);
            }
            while (queue.Count != 0)
            {
                var shape = queue.Dequeue();
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = shape.Ungroup();
                    foreach (var item in subRange)
                    {
                        queue.Enqueue(item as PowerPoint.Shape);
                    }
                }
                else if (!IsShape(shape))
                {
                    throw new CropLabException(TextCollection.CropToShapeText.ErrorMessageForSelectionNonShape);
                }
                else
                {
                    ungroupedShapeNames.Add(shape.Name);
                }
            }
            return currentSlide.Shapes.Range(ungroupedShapeNames.ToArray());
        }

        private static bool IsShapeForSelection(PowerPoint.ShapeRange shapeRange)
        {
            return (from PowerPoint.Shape shape in shapeRange select shape).All(IsShape);
        }

        private static bool IsShape(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape 
                || shape.Type == Office.MsoShapeType.msoFreeform 
                || shape.Type == Office.MsoShapeType.msoGroup;
        }
    }
}
