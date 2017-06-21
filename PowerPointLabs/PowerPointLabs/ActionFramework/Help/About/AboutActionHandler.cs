using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportActionRibbonId(TextCollection.AboutTag)]
    class AboutActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            MessageBox.Show(TextCollection.AboutInfo, TextCollection.AboutInfoTitle);
        }
    }
}
