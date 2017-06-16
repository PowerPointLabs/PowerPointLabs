using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.CropLab
{
    [ExportSupertipRibbonId(TextCollection.CropToAspectRatioTag + TextCollection.RibbonMenu)]
    class CropToAspectRatioSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.CropToAspectRatioSupertip;
        }
    }
}
