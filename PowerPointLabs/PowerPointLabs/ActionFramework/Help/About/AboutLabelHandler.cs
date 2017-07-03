using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportLabelRibbonId(TextCollection.AboutTag)]
    class AboutLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.AboutButtonLabel;
        }
    }
}
