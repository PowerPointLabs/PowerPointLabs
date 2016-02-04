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
        public void FitToHight(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio = false)
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

        public void FitToWidth(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio = false)
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

        public void FitToFill(PowerPoint.ShapeRange selectedShapes, bool isAspectRatio = false)
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
