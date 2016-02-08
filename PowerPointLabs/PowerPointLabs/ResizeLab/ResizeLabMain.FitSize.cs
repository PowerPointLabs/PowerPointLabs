using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using PowerPointLabs.Models;
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
        /// <param name="isAspectRatio"></param>
        public void FitToHight(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio)
        {
            if (isAspectRatio)
            {
                FitAspectRatioShapes(selectedShapes, Dimension.Height);
            }
            else
            {
                FitFreeShapes(selectedShapes, Dimension.Height);
            }
        }

        /// <summary>
        /// Fit selected shapes to the width of the slide.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="isAspectRatio"></param>
        public void FitToWidth(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio)
        {
            if (isAspectRatio)
            {
                FitAspectRatioShapes(selectedShapes, Dimension.Width);
            }
            else
            {
                FitFreeShapes(selectedShapes, Dimension.Width);
            }
        }

        /// <summary>
        /// Fit the selected shapes to fill the slide.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="isAspectRatio"></param>
        public void FitToFill(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio)
        {
            if (isAspectRatio)
            {
                FitAspectRatioShapes(selectedShapes, Dimension.HeightAndWidth);
            }
            else
            {
                FitFreeShapes(selectedShapes, Dimension.HeightAndWidth);
            }
        }

        /// <summary>
        /// Fit the selected shapes without aspect ratio according to the set dimension type.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="dimension"></param>
        private void FitFreeShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
        {
            try
            {
                var slideHight = PowerPointPresentation.Current.SlideHeight;
                var slideWidth = PowerPointPresentation.Current.SlideWidth;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];

                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.Height = slideHight;
                        shape.Top = 0;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.Width = slideWidth;
                        shape.Left = 0;
                    }
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "FitFreeShapes");
                throw;
            }
        }

        /// <summary>
        /// Fit the selected shapes with aspect ratio according to the set dimension type.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="dimension"></param>
        private void FitAspectRatioShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
        {
            try
            {
                var slideHight = PowerPointPresentation.Current.SlideHeight;
                var slideWidth = PowerPointPresentation.Current.SlideWidth;

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];

                    if (dimension == Dimension.Height)
                    {
                        FitToSlide.FitToHeight(shape, slideWidth, slideHight);
                    }
                    else if (dimension == Dimension.Width)
                    {
                        FitToSlide.FitToWidth(shape, slideWidth, slideHight);
                    }
                    else if (dimension == Dimension.HeightAndWidth)
                    {
                        FitToSlide.AutoFit(shape, slideWidth, slideHight);
                    }
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "FitAspectRatioShapes");
                throw;
            }
        }
    }
}
