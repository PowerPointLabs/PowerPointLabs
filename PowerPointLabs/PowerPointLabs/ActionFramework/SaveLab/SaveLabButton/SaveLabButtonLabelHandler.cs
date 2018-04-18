using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.SaveLab
{
    [ExportLabelRibbonId(SaveLabText.SavePresentationsButtonTag)]
    class SaveLabButtonLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return SaveLabText.SavePresentationsButtonLabel;
        }
    }
}
