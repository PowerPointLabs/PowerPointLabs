using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(TextCollection.EffectsLabColorizeTag)]
    class EffectsLabColorizeActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            PowerPointSlide curSlide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            if (ribbonId.Contains(TextCollection.ColorizeRemainderMenuId))
            {
                if (ribbonId.Contains(TextCollection.GrayScaleTag))
                {
                    EffectsLabColorize.GreyScaleRemainderEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.BlackWhiteTag))
                {
                    EffectsLabColorize.BlackWhiteRemainderEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.GothamTag))
                {
                    EffectsLabColorize.GothamRemainderEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.SepiaTag))
                {
                    EffectsLabColorize.SepiaRemainderEffect(curSlide, selection);
                }
                else
                {
                    Logger.Log(ribbonId + " does not exist!", Common.Logger.LogType.Error);
                }
            }
            else if (ribbonId.Contains(TextCollection.ColorizeBackgroundMenuId))
            {
                if (ribbonId.Contains(TextCollection.GrayScaleTag))
                {
                    EffectsLabColorize.GreyScaleBackgroundEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.BlackWhiteTag))
                {
                    EffectsLabColorize.BlackWhiteBackgroundEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.GothamTag))
                {
                    EffectsLabColorize.GothamBackgroundEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.SepiaTag))
                {
                    EffectsLabColorize.SepiaBackgroundEffect(curSlide, selection);
                }
                else
                {
                    Logger.Log(ribbonId + " does not exist!", Common.Logger.LogType.Error);
                }
            }
            else
            {
                Logger.Log(ribbonId + " does not exist!", Common.Logger.LogType.Error);
            }
        }
    }
}
