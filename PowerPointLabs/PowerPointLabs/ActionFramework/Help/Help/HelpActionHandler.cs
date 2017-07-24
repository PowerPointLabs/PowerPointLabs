using System.Diagnostics;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportActionRibbonId(HelpText.HelpTag)]
    class HelpActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Process.Start(CommonText.HelpDocumentUrl);
        }
    }
}
