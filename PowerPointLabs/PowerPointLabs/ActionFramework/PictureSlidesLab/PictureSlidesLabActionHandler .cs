using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.PictureSlidesLab.View;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportActionRibbonId(TextCollection.PictureSlidesLabTag)]
    class PictureSlidesLabActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            if (this.GetRibbonUi().PictureSlidesLabWindow == null || !this.GetRibbonUi().PictureSlidesLabWindow.IsOpen)
            {
                this.GetRibbonUi().PictureSlidesLabWindow = new PictureSlidesLabWindow();
                this.GetRibbonUi().PictureSlidesLabWindow.Show();
            }
            else
            {
                this.GetRibbonUi().PictureSlidesLabWindow.Activate();
            }
        }
    }
}
