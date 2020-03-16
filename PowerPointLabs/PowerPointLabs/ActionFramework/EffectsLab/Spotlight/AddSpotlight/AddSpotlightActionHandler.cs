using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(EffectsLabText.AddSpotlightTag)]
    class AddSpotlightActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            if (this.GetAddIn().Application.ActiveWindow.Selection.Type !=
                Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                return;
            }

            this.StartNewUndoEntry();
            PowerPointPresentation pres = this.GetCurrentPresentation();
            PowerPointSlide slide = this.GetCurrentSlide();

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                Spotlight.AddSpotlightEffect();
                return ClipboardUtil.ClipboardRestoreSuccess;
            }, pres, slide);
        }
    }
}
