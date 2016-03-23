using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using PowerPointLabs.Utils;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public partial class ResizeLabMain
    {

        private readonly ResizeLabErrorHandler _errorHandler;

        public ResizeLabMain()
        {
            _errorHandler = ResizeLabErrorHandler.InitializErrorHandler();
            SameDimensionAnchorType = SameDimensionAnchor.TopLeft;
        }

        private enum Dimension
        {
            Height,
            Width,
            HeightAndWidth
        }

        #region Validation

        /// <summary>
        /// Check if the selection is of shape type.
        /// </summary>
        /// <param name="selection"></param>
        /// <param name="handleError"></param>
        /// <param name="optionalParameters"></param>
        /// <returns></returns>
        internal bool IsSelectionValid(PowerPoint.Selection selection, bool handleError = true, string[] optionalParameters = null)
        {
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                if (handleError)
                {
                    _errorHandler.ProcessErrorCode(ResizeLabErrorHandler.ErrorCodeInvalidSelection, optionalParameters);
                }
                
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Check if the number of shape is more than one.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="minNoOfShapes"></param>
        /// <param name="handleError"></param>
        /// <param name="optionalParameters"></param>
        /// <returns></returns>
        private bool IsMoreThanOneShape(PowerPoint.ShapeRange selectedShapes, int minNoOfShapes, bool handleError = true, string[] optionalParameters = null)
        {
            if (selectedShapes.Count < minNoOfShapes)
            {
                if (handleError)
                {
                    _errorHandler.ProcessErrorCode(ResizeLabErrorHandler.ErrorCodeInvalidSelection, optionalParameters);
                }

                return false;
            }
            else
            {
                return true;
            }
        }

        #endregion

    }
}
