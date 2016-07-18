using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Supertip
{
    [ExportSupertipRibbonId("ColorsLabButton")]
    class ColorsLabSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId, string ribbonTag)
        {
            return TextCollection.ColorPickerButtonSupertip;
        }
    }
}
