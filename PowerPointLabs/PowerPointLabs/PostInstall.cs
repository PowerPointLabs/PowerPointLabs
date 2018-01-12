using Microsoft.VisualStudio.Tools.Applications.Deployment;

using PowerPointLabs.TextCollection;

namespace PowerPointLabs
{
    //Refer to:
    //http://msdn.microsoft.com/en-us/library/vstudio/dd465291(v=vs.100).aspx
    //Walkthrough: Copying a Document to the End User Computer after a ClickOnce Installation
    class PostInstall : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {
            var sourceFile = "";
            switch (Properties.Settings.Default.ReleaseType)
            {
                case "dev":
                    sourceFile = Properties.Settings.Default.DevAddr + CommonText.QuickTutorialFileName;
                    break;
                case "release":
                    sourceFile = Properties.Settings.Default.ReleaseAddr + CommonText.QuickTutorialFileName;
                    break;
            }

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                    try
                    {
                        if (sourceFile != "")
                        {
                            System.Diagnostics.Process.Start("POWERPNT", sourceFile);
                        }
                    }
                    catch
                    {
                        //MessageBox.Show("Can't open");
                    }
                    break;
                case AddInInstallationStatus.Update:
                    break;
                case AddInInstallationStatus.Uninstall:
                    break;
            }
        }
    }
}
