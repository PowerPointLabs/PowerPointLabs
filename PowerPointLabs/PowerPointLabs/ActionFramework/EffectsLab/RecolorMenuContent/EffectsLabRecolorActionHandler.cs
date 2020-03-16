using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.EffectsLab;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(EffectsLabText.RecolorTag)]
    class EffectsLabRecolorActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            PowerPointPresentation pres = this.GetCurrentPresentation();
            PowerPointSlide curSlide = this.GetCurrentSlide();
            Selection selection = this.GetCurrentSelection();

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                if (ribbonId.Contains(EffectsLabText.RecolorRemainderMenuId))
                {
                    if (ribbonId.Contains(EffectsLabText.GrayScaleTag))
                    {
                        EffectsLabRecolor.GrayScaleRemainderEffect(curSlide, selection);
                    }
                    else if (ribbonId.Contains(EffectsLabText.BlackWhiteTag))
                    {
                        EffectsLabRecolor.BlackWhiteRemainderEffect(curSlide, selection);
                    }
                    else if (ribbonId.Contains(EffectsLabText.GothamTag))
                    {
                        EffectsLabRecolor.GothamRemainderEffect(curSlide, selection);
                    }
                    else if (ribbonId.Contains(EffectsLabText.SepiaTag))
                    {
                        EffectsLabRecolor.SepiaRemainderEffect(curSlide, selection);
                    }
                    else
                    {
                        Logger.Log(ribbonId + " does not exist!", Common.Logger.LogType.Error);
                    }
                }
                else if (ribbonId.Contains(EffectsLabText.RecolorBackgroundMenuId))
                {
                    if (ribbonId.Contains(EffectsLabText.GrayScaleTag))
                    {
                        EffectsLabRecolor.GrayScaleBackgroundEffect(curSlide, selection);
                    }
                    else if (ribbonId.Contains(EffectsLabText.BlackWhiteTag))
                    {
                        EffectsLabRecolor.BlackWhiteBackgroundEffect(curSlide, selection);
                    }
                    else if (ribbonId.Contains(EffectsLabText.GothamTag))
                    {
                        EffectsLabRecolor.GothamBackgroundEffect(curSlide, selection);
                    }
                    else if (ribbonId.Contains(EffectsLabText.SepiaTag))
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
                return ClipboardUtil.ClipboardRestoreSuccess;
            }, pres, curSlide);
        }
    }
}
