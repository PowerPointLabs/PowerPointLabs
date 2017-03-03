using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CustomControls;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId("CropToSameButton")]
    class CropToSameActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop To Same Dimensions";

        protected override void ExecuteAction(string ribbonId)
        {
            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            if (!VerifyIsSelectionValid(this.GetCurrentSelection()))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 2, errorHandler);
                return;
            }
            ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            if (shapeRange.Count < 2)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypePicture, 2, errorHandler);
                return;
            }
            if (!IsPictureForSelection(shapeRange))
            {
                HandleErrorCodeIfRequired(CropLabErrorHandler.ErrorCodeSelectionMustBePicture, FeatureName, errorHandler);
                return;
            }
            CropToSame.CropSelection(shapeRange);
        }
    }
}
