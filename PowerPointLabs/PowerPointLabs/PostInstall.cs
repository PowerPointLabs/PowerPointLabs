using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using Microsoft.VisualStudio.Tools.Applications;
using System.IO;
using System.Windows.Forms;

namespace PowerPointLabs
{
    //Refer to:
    //http://msdn.microsoft.com/en-us/library/vstudio/dd465291(v=vs.100).aspx
    //Walkthrough: Copying a Document to the End User Computer after a ClickOnce Installation
    class PostInstall : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {
            string dataDirectory = @"PowerPointLabs Quick Tutorial.pptx";
            string sourcePath = args.AddInPath;
            string sourceFile = System.IO.Path.Combine(sourcePath, dataDirectory);

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                    try
                    {
                        System.Diagnostics.Process.Start(sourceFile);
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
