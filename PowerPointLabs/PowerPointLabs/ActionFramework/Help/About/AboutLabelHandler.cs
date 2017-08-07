using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportLabelRibbonId(HelpText.AboutTag)]
    class AboutLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return HelpText.AboutButtonLabel;
        }
    }
}
