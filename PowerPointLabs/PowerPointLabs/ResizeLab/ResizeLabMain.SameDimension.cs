using System;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// ResizeLabSameSize is the parital class of ResizeLabMain.
    /// It handles the resizing of the shapes to the same dimension 
    /// (e.g. height, width and both).
    /// </summary>
    internal partial class ResizeLabMain
    {
        /// <summary>
        /// Resize the selected shapes to the same height with the reference to
        /// first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameHeight(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.Height);
        }

        /// <summary>
        /// Resize the selected shapes to the same width with the reference to
        /// first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameWidth(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.Width);
        }

        /// <summary>
        /// Resize the selected shapes to the same size (width and height) with 
        /// the reference to first selected shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        public void ResizeToSameHeightAndWidth(PowerPoint.ShapeRange selectedShapes)
        {
            ResizeShapes(selectedShapes, Dimension.HeightAndWidth);
        }

        /// <summary>
        /// Resize the selected shapes according to the set dimension type.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="dimension"></param>
        private void ResizeShapes(PowerPoint.ShapeRange selectedShapes, Dimension dimension)
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
                    var shape = new PPShape(selectedShapes[i]);

                    if ((dimension == Dimension.Height) || (dimension == Dimension.HeightAndWidth))
                    {

                        shape.AbsoluteHeight = referenceHeight;
                    }

                    if ((dimension == Dimension.Width) || (dimension == Dimension.HeightAndWidth))
                    {
                        shape.AbsoluteWidth = referenceWidth;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ResizeShapes");
            }
        }

        /// <summary>
        /// Get the height of the reference shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <returns></returns>
        private float GetReferenceHeight(PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes.Count > 0)
            {
                return new PPShape(selectedShapes[1]).AbsoluteHeight;
            }
            return -1;
        }

        /// <summary>
        /// Get the width of the reference shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <returns></returns>
        private float GetReferenceWidth(PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes.Count > 0)
            {
                return new PPShape(selectedShapes[1]).AbsoluteWidth;
            }
            return -1;
        }
    }
}
