using System;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// FitToSlide is the partial class of ResizeLabMain.
    /// It handles fit to height, width and fill to the size of the slide.
    /// </summary>
    public partial class ResizeLabMain
    {
        // To be used for error handling
        internal const int Fit_MinNoOfShapesRequired = 1;
        internal const string Fit_FeatureName = "Fit To Slide";
        internal const string Fit_ShapeSupport = "object";
        internal static readonly string[] Fit_ErrorParameters =
        {
            Fit_FeatureName,
            Fit_MinNoOfShapesRequired.ToString(),
            Fit_ShapeSupport
        };

        /// <summary>
        /// Fit selected shapes to the height of the slide.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="slideHeight"></param>
        /// <param name="isAspectRatio"></param>
        /// <param name="slideWidth"></param>
        public void FitToHeight(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, bool isAspectRatio)
        {
            if (isAspectRatio)
            {
                FitAspectRatioShapes(selectedShapes, Dimension.Height, slideWidth, slideHeight);
            }
            else
            {
                FitFreeShapes(selectedShapes, Dimension.Height, slideWidth, slideHeight);
            }
        }

        /// <summary>
        /// Fit selected shapes to the width of the slide.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="slideHeight"></param>
        /// <param name="isAspectRatio"></param>
        /// <param name="slideWidth"></param>
        public void FitToWidth(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, bool isAspectRatio)
        {
            if (isAspectRatio)
            {
                FitAspectRatioShapes(selectedShapes, Dimension.Width, slideWidth, slideHeight);
            }
            else
            {
                FitFreeShapes(selectedShapes, Dimension.Width, slideWidth, slideHeight);
            }
        }

        /// <summary>
        /// Fit the selected shapes to fill the slide.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        /// <param name="isAspectRatio"></param>
        public void FitToFill(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, bool isAspectRatio)
        {
            if (isAspectRatio)
            {
                FitAspectRatioShapes(selectedShapes, Dimension.HeightAndWidth, slideWidth, slideHeight);
            }
            else
            {
                FitFreeShapes(selectedShapes, Dimension.HeightAndWidth, slideWidth, slideHeight);
            }
        }

        /// <summary>
        /// Fit the selected shapes without aspect ratio according to the set dimension type.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="dimension"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        private void FitFreeShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension, float slideWidth, float slideHeight)
        {
            try
            {
                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PPShape shape = new PPShape(selectedShapes[i]);

                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteHeight = slideHeight;
                        shape.VisualTop = 0;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteWidth = slideWidth;
                        shape.VisualLeft = 0;
                    }

                    if (dimension == Dimension.HeightAndWidth)
                    {
                        shape.VisualTop = 0;
                        shape.VisualLeft = 0;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "FitFreeShapes");
            }
        }

        /// <summary>
        /// Fit the selected shapes with aspect ratio according to the set dimension type.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="dimension"></param>
        /// <param name="slideWidth"></param>
        /// <param name="slideHeight"></param>
        private void FitAspectRatioShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension, float slideWidth, float slideHeight)
        {
            try
            {
                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];
                    System.Drawing.PointF anchorPoint = GetVisualAnchorPoint(new PPShape(shape, false));

                    if (dimension == Dimension.Height)
                    {
                        FitToSlide.FitToHeight(shape, slideWidth, slideHeight);

                        PPShape ppShape = new PPShape(shape, false);
                        AlignVisualShape(ppShape, anchorPoint);
                        ppShape.VisualTop = 0;
                    }
                    else if (dimension == Dimension.Width)
                    {
                        FitToSlide.FitToWidth(shape, slideWidth, slideHeight);

                        PPShape ppShape = new PPShape(shape, false);
                        AlignVisualShape(ppShape, anchorPoint);
                        ppShape.VisualLeft = 0;
                    }
                    else if (dimension == Dimension.HeightAndWidth)
                    {
                        FitToSlide.AutoFit(shape, slideWidth, slideHeight);
                    }
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "FitAspectRatioShapes");
            }
        }
    }
}
