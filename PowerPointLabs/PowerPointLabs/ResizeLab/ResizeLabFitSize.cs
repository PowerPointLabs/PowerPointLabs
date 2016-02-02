using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.Models;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    internal static partial class ResizeLabMain
    {
        public static void FitToHight(PowerPoint.ShapeRange selectedShapes)
        {
            FitShapes(selectedShapes, Dimension.Height);
        }

        public static void FitToWidth(PowerPoint.ShapeRange selectedShapes)
        {
            FitShapes(selectedShapes, Dimension.Width);
        }

        public static void Fill(PowerPoint.ShapeRange selectedShapes)
        {
            FitShapes(selectedShapes, Dimension.HeightAndWidth);
        }

        private static void FitShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
        {
            try
            {
                var slideHight = PowerPointPresentation.Current.SlideHeight;
                var slideWidth = PowerPointPresentation.Current.SlideWidth;

                for (int i = 1; i < selectedShapes.Count; i++)
                {
                    PowerPoint.Shape shape = selectedShapes[i];

                    switch (dimension)
                    {
                        case Dimension.Height:
                            FitToSlide.FitToHeight(shape, slideWidth, slideHight);
                            break;
                        case Dimension.Width:
                            FitToSlide.FitToWidth(shape, slideWidth, slideHight);
                            break;
                        case Dimension.HeightAndWidth:
                            FitToSlide.AutoFit(shape, slideWidth, slideHight);
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "FitShapes");
                throw;
            }
        }
    }
}
