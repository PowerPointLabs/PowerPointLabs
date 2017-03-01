using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropOutPaddingButton")]
    class CropOutPaddingActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop Out Padding";

        protected override void ExecuteAction(string ribbonId)
        {
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(CropLabUIControl.GetSharedInstance());
            var selection = this.GetCurrentSelection();

            if (!VerifyIsSelectionValid(selection))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }
            ShapeRange shapeRange = selection.ShapeRange;
            if (shapeRange.Count < 1)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }
            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, FeatureName, errorHandler);
                return;
            }

            try
            {
                CropOutPadding.Crop(selection);
            }
            catch (CropLabException e)
            {
                errorHandler.ProcessException(e, e.Message);
            }
        }
    }
}
