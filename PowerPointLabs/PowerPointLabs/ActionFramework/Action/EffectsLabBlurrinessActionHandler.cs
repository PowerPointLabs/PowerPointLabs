using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(TextCollection.EffectsLabBlurrinessTag)]
    class EffectsLabBlurrinessActionHandler : ActionHandler
    {
        private string feature;
        private Microsoft.Office.Interop.PowerPoint.Selection selection;
        private Models.PowerPointSlide slide;

        protected override void ExecuteAction(string ribbonId)
        {
            var isButton = false;
            int keywordIndex;

            if (ribbonId.Contains(TextCollection.DynamicMenuButtonId))
            {
                isButton = true;
                keywordIndex = ribbonId.IndexOf(TextCollection.DynamicMenuButtonId);
                feature = ribbonId.Substring(0, keywordIndex);
            }
            else
            {
                keywordIndex = ribbonId.IndexOf(TextCollection.DynamicMenuOptionId);
                feature = ribbonId.Substring(0, keywordIndex);
            }
            
            selection = this.GetCurrentSelection();
            slide = this.GetCurrentSlide();

            if (isButton)
            {
                if (!IsValidSelection())
                {
                    return;
                }

                var dialog = new EffectsLab.View.EffectsLabBlurrinessDialogBox(feature);
                dialog.SettingsHandler += PropertiesEdited;
                dialog.ShowDialog();
            }
            else
            {
                var startIndex = keywordIndex + TextCollection.DynamicMenuOptionId.Length;
                var percentage = int.Parse(ribbonId.Substring(startIndex, ribbonId.Length - startIndex));
                ExecuteBlurAction(percentage);
            }
        }

        private bool IsValidSelection()
        {
            if (EffectsLab.EffectsLabBlurSelected.IsValidSelection(selection)
                && EffectsLab.EffectsLabBlurSelected.IsValidShapeRange(selection.ShapeRange))
            {
                return true;
            }

            return false;
        }

        private void PropertiesEdited(int percentage, bool isTint)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    EffectsLab.EffectsLabBlurSelected.IsTintSelected = isTint;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    EffectsLab.EffectsLabBlurSelected.IsTintRemainder = isTint;
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    EffectsLab.EffectsLabBlurSelected.IsTintBackground = isTint;
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }

            this.GetRibbonUi().RefreshRibbonControl(feature + TextCollection.DynamicMenuCheckBoxId);

            ExecuteBlurAction(percentage);
        }

        private void ExecuteBlurAction(int percentage)
        {
            switch (feature)
            {
                case TextCollection.EffectsLabBlurrinessFeatureSelected:
                    this.StartNewUndoEntry();
                    EffectsLab.EffectsLabBlurSelected.BlurSelected(slide, selection, percentage);
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureRemainder:
                    this.GetRibbonUi().BlurRemainderEffectClick(percentage);
                    break;
                case TextCollection.EffectsLabBlurrinessFeatureBackground:
                    this.GetRibbonUi().BlurBackgroundEffectClick(percentage);
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }
    }
}
