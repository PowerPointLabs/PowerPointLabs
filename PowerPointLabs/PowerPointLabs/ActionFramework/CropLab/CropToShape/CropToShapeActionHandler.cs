using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CustomControls;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(CropLabText.CropToShapeTag)]
    class CropToShapeActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop To Shape";

        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            if (!ShapeUtil.IsSelectionShape(this.GetCurrentSelection()))
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypeShape, 1, errorHandler);
                return;
            }
            ShapeRange shapeRange = this.GetCurrentSelection().ShapeRange;
            if (shapeRange.Count < 1)
            {
                HandleInvalidSelectionError(CropLabErrorHandler.ErrorCodeSelectionIsInvalid, FeatureName, CropLabErrorHandler.SelectionTypeShape, 1, errorHandler);
                return;
            }
            if (!ShapeUtil.IsAllShape(shapeRange))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeSelectionMustBeShape, FeatureName, errorHandler);
                return;
            }
            try
            {
                CropToShape.Crop(this.GetCurrentSlide(), this.GetCurrentSelection());
            }
            catch (CropLabException)
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeUndefined, FeatureName, errorHandler);
            }
        }
    }
}
