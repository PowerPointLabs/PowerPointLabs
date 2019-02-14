using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportEnabledRibbonId(PictureSlidesLabText.PaneTag)]
    class PictureSlidesEnabledHandler : EnabledHandler
    {
        protected override bool GetEnabled(string ribbonId)
        {
            return PowerPointLabs.PictureSlidesLab.PictureSlidesLab.IsPictureSlidesEnabled;
        }
    }
}