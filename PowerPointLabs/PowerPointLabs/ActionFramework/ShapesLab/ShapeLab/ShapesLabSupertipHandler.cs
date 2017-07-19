using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportSupertipRibbonId(TextCollection.ShapesLabTag)]
    class ShapesLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.CustomShapeButtonSupertip;
        }
    }
}
