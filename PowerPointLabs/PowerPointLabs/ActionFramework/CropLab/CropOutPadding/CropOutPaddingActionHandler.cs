using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CustomControls;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;


namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(CropLabText.CropOutPaddingTag)]
    class CropOutPaddingActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop Out Padding";

        protected override void ExecuteAction(string ribbonId)
        {
            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            Selection selection = this.GetCurrentSelection();

            if (!ShapeUtil.IsSelectionShape(selection))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }

            ShapeRange shapeRange = ShapeUtil.GetShapeRange(selection);

            if (shapeRange.Count < 1)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 1, errorHandler);
                return;
            }
            if (!ShapeUtil.IsAllPicture(shapeRange))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, FeatureName, errorHandler);
                return;
            }

            try
            {
                ShapeRange result = CropOutPadding.Crop(shapeRange);
                result?.Select();
            }
            catch (CropLabException e)
            {
                HandleCropLabException(e, FeatureName, errorHandler);
            }
        }
    }
}
