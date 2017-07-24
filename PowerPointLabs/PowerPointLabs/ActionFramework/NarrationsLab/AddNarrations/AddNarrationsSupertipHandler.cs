using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportSupertipRibbonId(NarrationsLabText.AddNarrationsTag)]
    class AddNarrationsSupertipHandler : SupertipHandler
    {
        protected override string GetSupertip(string ribbonId)
        {
            return NarrationsLabText.AddNarrationsButtonSupertip;
        }
    }
}
