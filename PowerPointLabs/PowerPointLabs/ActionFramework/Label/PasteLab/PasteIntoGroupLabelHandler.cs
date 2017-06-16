using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        TextCollection.PasteIntoGroupId + TextCollection.MenuGroup,
        TextCollection.PasteIntoGroupId + TextCollection.RibbonButton)]
    class PasteIntoGroupLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteIntoGroup;
        }
    }
}