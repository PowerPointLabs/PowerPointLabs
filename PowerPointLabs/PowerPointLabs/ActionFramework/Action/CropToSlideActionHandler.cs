using System;
using System.Drawing;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CropLab;
using Office = Microsoft.Office.Core;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropToSlideButton")]
    class CropToSlideActionHandler : CropLabActionHandler
    {
        private static readonly string ShapePicture = Path.GetTempPath() + @"\shape.png";
        private static readonly string FeatureName = "Crop To Slide";

        protected override void ExecuteAction(string ribbonId)
        {
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(CropLabUIControl.GetSharedInstance());
            if (!VerifyIsSelectionValid(this.GetCurrentSelection()))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }
            ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            if (shapeRange.Count < 1)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }
            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, FeatureName, errorHandler);
                return;
            }
            foreach (Shape shape in shapeRange)
            {
                Shape toRotate = shape;
                if (shape.Rotation != 0)
                {
                    RectangleF location = GetAbsoluteBounds(shape);
                    Utils.Graphics.ExportShape(shape, ShapePicture);
                    var newShape = this.GetCurrentSlide().Shapes.AddPicture(ShapePicture,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        location.Left, location.Top, location.Width, location.Height);
                    toRotate = newShape;
                    toRotate.Name = shape.Name;
                    shape.Delete();

                }
                float slideWidth = this.GetCurrentPresentation().SlideWidth;
                float slideHeight = this.GetCurrentPresentation().SlideHeight;
                RectangleF cropArea = GetCropArea(toRotate, slideWidth, slideHeight);
                toRotate.PictureFormat.Crop.ShapeHeight = cropArea.Height;
                toRotate.PictureFormat.Crop.ShapeWidth = cropArea.Width;
                toRotate.PictureFormat.Crop.ShapeLeft = cropArea.Left;
                toRotate.PictureFormat.Crop.ShapeTop = cropArea.Top;
            }
        }

        private static RectangleF GetAbsoluteBounds(Shape shape)
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

        private static RectangleF GetCropArea(Shape shape, float slideWidth, float slideHeight)
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
