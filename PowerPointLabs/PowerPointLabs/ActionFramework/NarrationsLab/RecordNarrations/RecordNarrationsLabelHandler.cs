using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportLabelRibbonId(NarrationsLabText.RecordNarrationsTag)]
    class RecordNarrationsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return NarrationsLabText.RecordNarrationsButtonLabel;
        }
    }
}
