using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.TooltipsLab
{
    [ExportSupertipRibbonId(TooltipsLabText.CreateTriggerTag)]
    class CreateTriggerSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return TooltipsLabText.CreateTriggerButtonSupertip;
        }
    }
}