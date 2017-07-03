using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ColorsLab
{
    [ExportSupertipRibbonId(TextCollection.ColorsLabTag)]
    class ColorsLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TextCollection.ColorsLabMenuSupertip;
        }
    }
}
