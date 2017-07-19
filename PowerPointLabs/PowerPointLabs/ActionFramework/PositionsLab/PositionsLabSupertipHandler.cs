using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.PositionsLab
{
    [ExportSupertipRibbonId(TextCollection.PositionsLabTag)]
    class PositionsLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.PositionsLabMenuSupertip;
        }
    }
}
