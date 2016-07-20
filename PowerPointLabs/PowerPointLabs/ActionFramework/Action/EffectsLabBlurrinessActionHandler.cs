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

        protected override void ExecuteAction(string ribbonId, string ribbonTag)
        {
            var isButton = false;
            int keywordIndex;

            if (ribbonId.Contains("Button"))
            {
                isButton = true;
                keywordIndex = ribbonId.IndexOf("Button");
                feature = ribbonId.Substring(0, keywordIndex);
            }
            else
            {
                keywordIndex = ribbonId.IndexOf("Option");
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

                var dialog = new Views.EffectsLabBlurrinessDialogBox();
                dialog.SettingsHandler += PropertiesEdited;
                dialog.ShowDialog();
            }
            else
            {
                var startIndex = keywordIndex + 6;
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

        private void PropertiesEdited(int percentage, bool hasOverlay)
        {
            EffectsLab.EffectsLabBlurSelected.HasOverlay = hasOverlay;
            var ribbon = this.GetRibbonUi();
            ribbon.RefreshRibbonControl("EffectsLabBlurSelectedCheckBox");
            ribbon.RefreshRibbonControl("EffectsLabBlurRemainderCheckBox");
            ribbon.RefreshRibbonControl("EffectsLabBlurBackgroundCheckBox");

            ExecuteBlurAction(percentage);
        }

        private void ExecuteBlurAction(int percentage)
        {
            switch (feature)
            {
                case "EffectsLabBlurSelected":
                    this.StartNewUndoEntry();
                    EffectsLab.EffectsLabBlurSelected.BlurSelected(slide, selection, percentage);
                    break;
                case "EffectsLabBlurRemainder":
                    this.GetRibbonUi().BlurRemainderEffectClick(percentage);
                    break;
                case "EffectsLabBlurBackground":
                    this.GetRibbonUi().BlurBackgroundEffectClick(percentage);
                    break;
                default:
                    throw new System.Exception("Invalid feature");
            }
        }
    }
}
