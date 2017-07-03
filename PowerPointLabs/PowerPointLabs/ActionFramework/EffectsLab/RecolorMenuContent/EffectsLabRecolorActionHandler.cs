using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.Models;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(TextCollection.EffectsLabRecolorTag)]
    class EffectsLabRecolorActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            PowerPointSlide curSlide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            if (ribbonId.Contains(TextCollection.RecolorRemainderMenuId))
            {
                if (ribbonId.Contains(TextCollection.GrayScaleTag))
                {
                    EffectsLabRecolor.GreyScaleRemainderEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.BlackWhiteTag))
                {
                    EffectsLabRecolor.BlackWhiteRemainderEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.GothamTag))
                {
                    EffectsLabRecolor.GothamRemainderEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.SepiaTag))
                {
                    EffectsLabRecolor.SepiaRemainderEffect(curSlide, selection);
                }
                else
                {
                    Logger.Log(ribbonId + " does not exist!", Common.Logger.LogType.Error);
                }
            }
            else if (ribbonId.Contains(TextCollection.RecolorBackgroundMenuId))
            {
                if (ribbonId.Contains(TextCollection.GrayScaleTag))
                {
                    EffectsLabRecolor.GreyScaleBackgroundEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.BlackWhiteTag))
                {
                    EffectsLabRecolor.BlackWhiteBackgroundEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.GothamTag))
                {
                    EffectsLabRecolor.GothamBackgroundEffect(curSlide, selection);
                }
                else if (ribbonId.Contains(TextCollection.SepiaTag))
                {
                    EffectsLabRecolor.SepiaBackgroundEffect(curSlide, selection);
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
