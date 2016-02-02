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
        public static void ResizeToSameHeight(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.Height);
        }

        public static void ResizeToSameWidth(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.Width);
        }

        public static void ResizeToSameHeightAndWidth(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.HeightAndWidth);
        }

        #region General

        private static void ResizeShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
        {
            try
            {
                var referenceHeight = GetReferenceHeight(selectedShapes);
                var referenceWidth = GetReferenceWidth(selectedShapes);

                if (!IsMoreThanOneShape(selectedShapes) || (referenceHeight < 0) || (referenceWidth < 0))
                {
                    return;
                }

                for (int i = 1; i <= selectedShapes.Count; i++)
                {
                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {
                        selectedShapes[i].Height = referenceHeight;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
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
