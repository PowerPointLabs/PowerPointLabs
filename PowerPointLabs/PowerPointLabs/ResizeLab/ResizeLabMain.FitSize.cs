using System;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// ResizeLabFitSize is the parital class of ResizeLabMain.
    /// It handles fit to height, width and fill to the size of the slide.
    /// </summary>
    internal partial class ResizeLabMain
    {
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
                    var shape = new PPShape(selectedShapes[i]);

                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteHeight = slideHeight;
                        shape.Top = 0;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteWidth = slideWidth;
                        shape.Left = 0;
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

                    if (dimension == Dimension.Height)
                    {
                        FitToSlide.FitToHeight(shape, slideWidth, slideHeight);
                    }
                    else if (dimension == Dimension.Width)
                    {
                        FitToSlide.FitToWidth(shape, slideWidth, slideHeight);
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
