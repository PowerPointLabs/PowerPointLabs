using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip.CropLab
{
    [ExportSupertipRibbonId(TextCollection.CropToShapeTag + TextCollection.RibbonButton)]
    class CropToShapeSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.MoveCropShapeButtonSupertip;
        }
    }
}
