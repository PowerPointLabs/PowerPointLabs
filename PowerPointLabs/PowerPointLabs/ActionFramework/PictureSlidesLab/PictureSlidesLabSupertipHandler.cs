using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportSupertipRibbonId(PictureSlidesLabText.PaneTag)]
    class PictureSlidesLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return PictureSlidesLabText.RibbonMenuSupertip;
        }
    }
}
