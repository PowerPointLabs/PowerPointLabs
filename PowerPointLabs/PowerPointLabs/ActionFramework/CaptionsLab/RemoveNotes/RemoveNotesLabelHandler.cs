using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportLabelRibbonId(TextCollection.RemoveNotesTag)]
    class RemoveNotesLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.RemoveAllNotesButtonLabel;
        }
    }
}
