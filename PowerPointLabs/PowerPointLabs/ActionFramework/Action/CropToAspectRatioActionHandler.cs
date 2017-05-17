using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CustomControls;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(TextCollection.CropToAspectRatioTag)]
    class CropToAspectRatioActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop To Aspect Ratio";

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

            if (ribbonId.Contains(TextCollection.DynamicMenuButtonId))
            {
                var dialog = new CustomAspectRatioDialog();
                dialog.SettingsHandler += ExecuteCropToAspectRatio;
                dialog.ShowDialog();
            }
            else if (ribbonId.Contains(TextCollection.DynamicMenuOptionId))
            {
                int optionRawStringStartIndex = ribbonId.LastIndexOf(TextCollection.DynamicMenuButtonId) +
                                                TextCollection.DynamicMenuOptionId.Length;
                string optionRawString = ribbonId.Substring(optionRawStringStartIndex).Replace('_', ':');
                ExecuteCropToAspectRatio(optionRawString);
            }
        }

        private void ExecuteCropToAspectRatio(string aspectRatioRawString)
        {
            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            var selection = this.GetCurrentSelection();

            float aspectRatioWidth = 0.0f;
            float aspectRatioHeight = 0.0f;
            if (!TryParseAspectRatio(aspectRatioRawString, out aspectRatioWidth, out aspectRatioHeight))
            {
                HandleErrorCode(CropLabErrorHandler.ErrorCodeAspectRatioIsInvalid, FeatureName, errorHandler);
                return;
            }
            float aspectRatio = aspectRatioWidth / aspectRatioHeight;

            try
            {
                CropToAspectRatio.Crop(selection, aspectRatio);
            }
            catch (CropLabException e)
            {
                HandleCropLabException(e, FeatureName, errorHandler);
            }
        }
    }
}
