using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    internal static partial class ResizeLabMain
    {
        private enum Dimension
        {
            Height,
            Width,
            HeightAndWidth
        }

        public static void ResizeToSameHeight()
        {
            ResizeShapes(Dimension.Height);
        }

        public static void ResizeToSameWidth()
        {
            ResizeShapes(Dimension.Width);
        }

        public static void ResizeToSameHeightAndWidth()
        {
            ResizeShapes(Dimension.HeightAndWidth);
        }

        #region General

        private static void ResizeShapes(Dimension dimensionType)
        {
            try
            {
                var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    // Show error message
                    return;
                }

                var selectedShapes = selection.ShapeRange;
                var referenceHeight = GetReferenceHeight(selectedShapes);
                var referenceWidth = GetReferenceWidth(selectedShapes);

                if ((selectedShapes.Count < 2) || (referenceHeight < 0) || (referenceWidth < 0))
                {
                    // Show error message
                    return;
                }

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    if ((dimensionType == Dimension.Height) || (dimensionType == Dimension.HeightAndWidth))
                    {
                        selectedShapes[i].Height = referenceHeight;
                    }

                    if ((dimensionType == Dimension.Width) || (dimensionType == Dimension.HeightAndWidth))
                    {
                        selectedShapes[i].Width = referenceWidth;
                    }
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ResizeShapes");
                throw;
            }
            
        }

        private static float GetReferenceHeight(PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes.Count > 0)
            {
                return selectedShapes[1].Height;
            }
            return -1;
        }

        private static float GetReferenceWidth(PowerPoint.ShapeRange selectShapes)
        {
            if (selectShapes.Count > 0)
            {
                return selectShapes[1].Width;
            }
            return -1;
        }

        #endregion
    }
}
