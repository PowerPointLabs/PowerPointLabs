using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.ZoomLab
{
    [ExportLabelRibbonId(TextCollection.RecordNarrationsTag)]
    class RecordNarrationsLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.RecordNarrationsButtonLabel;
        }
    }
}
