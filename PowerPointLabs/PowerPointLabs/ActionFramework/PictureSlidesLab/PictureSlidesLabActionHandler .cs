using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.PictureSlidesLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportActionRibbonId(PictureSlidesLabText.PaneTag)]
    class PictureSlidesLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            if (this.GetRibbonUi().PictureSlidesLabWindow == null || !this.GetRibbonUi().PictureSlidesLabWindow.IsOpen)
            {
                this.GetRibbonUi().PictureSlidesLabWindow = new PictureSlidesLabWindow();
                this.GetRibbonUi().PictureSlidesLabWindow.ShowThematicDialog(false);
            }
            else
            {
                this.GetRibbonUi().PictureSlidesLabWindow.Activate();
            }
        }
    }
}
