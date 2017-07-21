using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportLabelRibbonId(TextCollection1.HelpTag)]
    class HelpLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return HelpText.UserGuideButtonLabel;
        }
    }
}
