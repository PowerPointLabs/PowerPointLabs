using System.Diagnostics;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ActionFramework.Help
{
    [ExportActionRibbonId(TextCollection.TutorialTag)]
    class TutorialActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            string sourceFile = "";
            switch (Properties.Settings.Default.ReleaseType)
            {
                case "dev":
                    sourceFile = Properties.Settings.Default.DevAddr + TextCollection.QuickTutorialFileName;
                    break;
                case "release":
                    sourceFile = Properties.Settings.Default.ReleaseAddr + TextCollection.QuickTutorialFileName;
                    break;
            }

            try
            {
                if (sourceFile != "")
                {
                    Process.Start("POWERPNT", sourceFile);
                }
            }
            catch
            {
                Logger.Log("TutorialButtonClick: Failed to open tutorial file!", ActionFramework.Common.Logger.LogType.Error);
            }
        }
    }
}
