using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.CropLab;
using PowerPointLabs.CropLab.Views;
using PowerPointLabs.CustomControls;
using PowerPointLabs.Extensions;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.CropLab
{
    [ExportActionRibbonId(CropLabText.CropToAspectRatioTag + CommonText.RibbonMenu)]
    class CropToAspectRatioActionHandler : CropLabActionHandler
    {
        private static readonly string FeatureName = "Crop To Aspect Ratio";

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
            ShapeRange shapeRange = selection.ShapeRange;
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

            if (ribbonId.Contains(CommonText.DynamicMenuButtonId))
            {
                CustomAspectRatioDialogBox dialog = new CustomAspectRatioDialogBox(shapeRange[1]);
                dialog.DialogConfirmedHandler += ExecuteCropToAspectRatio;
                Globals.ThisAddIn.ColorThemeChanged += dialog.UpdateColors;
                dialog.ShowDialog();
                Globals.ThisAddIn.ColorThemeChanged -= dialog.UpdateColors;
            }
            else if (ribbonId.Contains(CommonText.DynamicMenuOptionId))
            {
                int optionRawStringStartIndex = ribbonId.LastIndexOf(CommonText.DynamicMenuButtonId) +
                                                CommonText.DynamicMenuOptionId.Length;
                string optionRawString = ribbonId.Substring(optionRawStringStartIndex).Replace('_', ':');
                ExecuteCropToAspectRatio(optionRawString);
            }
        }

        private void ExecuteCropToAspectRatio(string aspectRatioRawString)
        {
            IMessageService cropLabMessageService = MessageServiceFactory.GetCropLabMessageService();
            CropLabErrorHandler errorHandler = CropLabErrorHandler.InitializeErrorHandler(cropLabMessageService);
            Selection selection = this.GetCurrentSelection();

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
