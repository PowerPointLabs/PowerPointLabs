using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.CaptionsLab
{
    [ExportLabelRibbonId(CaptionsLabText.RemoveNotesTag)]
    class RemoveNotesLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return CaptionsLabText.RemoveAllNotesButtonLabel;
        }
    }
}
