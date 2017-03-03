using System;
using System.Drawing;
using System.IO;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.CropLab
{
    public class CropToSlide
    {
        private static readonly string ShapePicture = Path.GetTempPath() + @"\shape.png";

        public static void CropSelection(PowerPoint.ShapeRange shapeRange, PowerPointSlide currentSlide, float slideWidth, float slideHeight)
        {
            foreach (PowerPoint.Shape shape in shapeRange)
            {
                PowerPoint.Shape toRotate = shape;
                if (shape.Rotation != 0)
                {
                    RectangleF location = GetAbsoluteBounds(shape);
                    Utils.Graphics.ExportShape(shape, ShapePicture);
                    var newShape = currentSlide.Shapes.AddPicture(ShapePicture,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        location.Left, location.Top, location.Width, location.Height);
                    toRotate = newShape;
                    toRotate.Name = shape.Name;
                    shape.Delete();

                }
                RectangleF cropArea = GetCropArea(toRotate, slideWidth, slideHeight);
                toRotate.PictureFormat.Crop.ShapeHeight = cropArea.Height;
                toRotate.PictureFormat.Crop.ShapeWidth = cropArea.Width;
                toRotate.PictureFormat.Crop.ShapeLeft = cropArea.Left;
                toRotate.PictureFormat.Crop.ShapeTop = cropArea.Top;
            }
        }

        private static RectangleF GetAbsoluteBounds(PowerPoint.Shape shape)
        {
            float rotation = (float)Utils.Graphics.DegreeToRadian(shape.Rotation);
            PointF[] corners = new PointF[]
            {
                new PointF(-shape.Width / 2, -shape.Height / 2),
                new PointF(shape.Width / 2, -shape.Height / 2),
                new PointF(-shape.Width / 2, shape.Height / 2),
                new PointF(shape.Width / 2, shape.Height / 2)
            };
            float minX = float.MaxValue;
            float minY = float.MaxValue;
            float maxX = float.MinValue;
            float maxY = float.MinValue;
            for (int i = 0; i < corners.Length; i++)
            {
                PointF rotated = RotatePoint(corners[i], rotation);
                minX = Math.Min(rotated.X, minX);
                minY = Math.Min(rotated.Y, minY);
                maxX = Math.Max(rotated.X, maxX);
                maxY = Math.Max(rotated.Y, maxY);
            }
            return new RectangleF(shape.Left + shape.Width / 2 + minX, shape.Top + shape.Height / 2 + minY,
                                  maxX - minX, maxY - minY);
        }

        private static PointF RotatePoint(PointF point, float theta)
        {
            return new PointF((float)(point.X * Math.Cos(theta) - point.Y * Math.Sin(theta)),
                            (float)(point.X * Math.Sin(theta) + point.Y * Math.Cos(theta)));
        }

        private static RectangleF GetCropArea(PowerPoint.Shape shape, float slideWidth, float slideHeight)
        {
            float cropTop = Math.Max(0, shape.Top);
            float cropLeft = Math.Max(0, shape.Left);
            float cropHeight = shape.Height - Math.Max(0, -shape.Top);
            float cropWidth = shape.Width - Math.Max(0, -shape.Left);

            cropHeight = Math.Min(slideHeight - cropTop, cropHeight);
            cropWidth = Math.Min(slideWidth - cropLeft, cropWidth);

            return new RectangleF(cropLeft, cropTop, cropWidth, cropHeight);
        }
    }
}
