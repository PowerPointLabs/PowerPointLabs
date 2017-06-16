using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.CropLab
{
    [ExportSupertipRibbonId(TextCollection.CropToSlideTag + TextCollection.RibbonButton)]
    class CropToSlideSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.CropToSlideButtonSupertip;
        }
    }
}
