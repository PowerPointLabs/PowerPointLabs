using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        TextCollection.PasteIntoGroupTag + TextCollection.MenuGroup,
        TextCollection.PasteIntoGroupTag + TextCollection.RibbonButton)]
    class PasteIntoGroupLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteIntoGroup;
        }
    }
}