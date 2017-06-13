using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.HelpMenu
{
    [ExportLabelRibbonId("TutorialButton")]
    class TutorialLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.TutorialButtonLabel;
        }
    }
}
