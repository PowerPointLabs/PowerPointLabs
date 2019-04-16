using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public partial class ResizeLabMain
    {

        private readonly ResizeLabErrorHandler _errorHandler;

        public ResizeLabMain()
        {
            _errorHandler = ResizeLabErrorHandler.InitializeErrorHandler();
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
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                if (handleError)
                {
                    _errorHandler.ProcessErrorCode(ResizeLabErrorHandler.ErrorCodeInvalidSelection, optionalParameters);
                }

                return false;
            }
            return true;
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
            return true;
        }

        #endregion

        /// <summary>
        /// Get the height of the reference shape.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <returns></returns>
        private float GetReferenceHeight(PowerPoint.ShapeRange selectedShapes)
        {
            if (selectedShapes.Count <= 0)
            {
                return -1;
            }

            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    return new PPShape(selectedShapes[1]).AbsoluteHeight;
                case ResizeBy.Actual:
                    return new PPShape(selectedShapes[1], false).ShapeHeight;
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
            if (selectedShapes.Count <= 0)
            {
                return -1;
            }

            switch (ResizeType)
            {
                case ResizeBy.Visual:
                    return new PPShape(selectedShapes[1]).AbsoluteWidth;
                case ResizeBy.Actual:
                    return new PPShape(selectedShapes[1], false).ShapeWidth;
            }
            return -1;
        }
    }
}
