using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportLabelRibbonId(NarrationsLabText.AddNarrationsTag)]
    class AddNarrationsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return NarrationsLabText.AddNarrationsButtonLabel;
        }
    }
}
