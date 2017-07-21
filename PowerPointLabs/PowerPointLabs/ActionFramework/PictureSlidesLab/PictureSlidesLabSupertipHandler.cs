using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PictureSlidesLab
{
    [ExportSupertipRibbonId(TextCollection1.PictureSlidesLabTag)]
    class PictureSlidesLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection1.PictureSlidesLabMenuSupertip;
        }
    }
}
