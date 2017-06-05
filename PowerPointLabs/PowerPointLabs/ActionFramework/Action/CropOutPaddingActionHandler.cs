using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CustomControls;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropOutPaddingButton")]
    class CropOutPaddingActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop Out Padding";

        protected override void ExecuteAction(string ribbonId)
        {
            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            var selection = this.GetCurrentSelection();

            if (!IsSelectionShapes(selection))
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
            if (!IsAllPicture(shapeRange))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, FeatureName, errorHandler);
                return;
            }

            try
            {
                CropOutPadding.Crop(selection);
            }
            catch (CropLabException e)
            {
                HandleCropLabException(e, FeatureName, errorHandler);
            }
        }
    }
}
