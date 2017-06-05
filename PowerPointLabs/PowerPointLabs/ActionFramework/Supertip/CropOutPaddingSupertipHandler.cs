using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip
{
    [ExportSupertipRibbonId("CropOutPaddingButton")]
    class CropOutPaddingSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.CropOutPaddingSupertip;
        }
    }
}
