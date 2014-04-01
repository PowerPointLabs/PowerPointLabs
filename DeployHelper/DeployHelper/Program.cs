using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

//
//  DeployHelper Class
//  ------------------
//  Simply double click the .exe file to patch PowerPointLabs, 
//  to make it support PostInstall event (e.g. open tutorial file after install pptlabs)
//
//  HOW TO USE
//
//  1. Fill in DeployHelper.conf
//  - Mage is a component provided by Visual Studio
//  - Key is inside PowerPointLabs project
//
//  2. Copy both DeployHelper.exe and DeployHelper.conf to the publish folder
//
//  3. Run DeployHelper.exe
//

namespace DeployHelper
{
    class Program
    {
        const string ERROR_NO_CONFIG = "Can't Find Config.";
        const string ERROR_NO_VSTO = "Can't Find VSTO.";
        const string ERROR_NO_MANIFEST = "Can't Find Manifest For This Version";
        const string ERROR_ALREADY_PATCHED = "Already Patched.";
        const string ERROR_INVALID_KEY_OR_MAGE_DIR = "Invalid Mage or Key Directory.";
        static void DisplayWarning(string content)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(content);
            Console.ResetColor();
            Console.ReadKey();
        }

        static void DisplayDone()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Patched.");
            Console.ResetColor();
            Console.ReadKey();
        }

        static string AddQuote(string dir)
        {
            return "\"" + dir + "\"";
        }

        static void Main(string[] args)
        {
            //Reference on What It Does
            //http://msdn.microsoft.com/en-us/library/vstudio/dd465291(v=vs.100).aspx


            //*****************
            //***Read Config***
            //*****************
            string currentDirectory = System.Environment.CurrentDirectory;
            string vstoDirectory = currentDirectory + @"\PowerPointLabs.vsto";
            string configDirectory = currentDirectory + @"\DeployHelper.conf";
            string[] configContent;

            //get keyDir and mageDir from config file
            //use them to re-sign
            try
            {
                configContent = System.IO.File.ReadAllLines(configDirectory);
            }
            catch
            {
                DisplayWarning(ERROR_NO_CONFIG);
                return;
            }
            string mageDirectory = configContent[1];
            string keyDirectory = configContent[3];

            //get version from VSTO file
            XmlDocument currentVsto;
            try
            {
                currentVsto = new XmlDocument();
                currentVsto.Load(vstoDirectory);
            }
            catch
            {
                DisplayWarning(ERROR_NO_VSTO);
                return;
            }
            var vstoNode = currentVsto.GetElementsByTagName("assemblyIdentity")[0];
            string version = vstoNode.Attributes["version"].Value;

            string[] versionDetails = version.Split('.');
            string versionMajor = versionDetails[0];
            string versionMinor = versionDetails[1];
            string versionBuild = versionDetails[2];
            string versionRevision = versionDetails[3];

            string buildDirectory = currentDirectory + @"\Application Files\PowerPointLabs_"
                + versionMajor + "_" + versionMinor + "_" + versionBuild + "_" + versionRevision;
            string manifestDirectory = buildDirectory + @"\PowerPointLabs.dll.manifest";
            string destVstoDirectory = buildDirectory + @"\PowerPointLabs.vsto";

            //*********************
            //***Modify Manifest***
            //*********************
            XmlDocument doc = new XmlDocument();
            XmlDocument docBackup = new XmlDocument();
            try
            {
                doc.Load(manifestDirectory);
                docBackup.Load(manifestDirectory);
            }
            catch
            {
                DisplayWarning(ERROR_NO_MANIFEST);
                return;
            }
            //If Not Patched
            if (doc.GetElementsByTagName("vstav3:postAction").Count == 0)
            {
                //Patch Content for Manifest
                //************************************************************
                //<vstav3:postActions>
                //  <vstav3:postAction>
                //      <vstav3:entryPoint class="PowerPointLabs.PostInstall">
                //          <assemblyIdentity 
                //          name="PostInstall" 
                //          version="{$version}" 
                //          language="neutral" 
                //          processorArchitecture="msil"/>
                //      </vstav3:entryPoint>
                //      <vstav3:postActionData>
                //      </vstav3:postActionData>
                //  </vstav3:postAction>
                //</vstav3:postActions>
                //************************************************************
                XmlNode addInNode = doc.GetElementsByTagName("vstav3:addIn")[0];
                XmlNode updateNode = doc.GetElementsByTagName("vstav3:update")[0];
                XmlElement postActionsNode = doc.CreateElement("vstav3", "postActions", "urn:schemas-microsoft-com:vsta.v3");
                XmlElement postActionNode = doc.CreateElement("vstav3", "postAction", "urn:schemas-microsoft-com:vsta.v3");
                XmlElement entryPointNode = doc.CreateElement("vstav3", "entryPoint", "urn:schemas-microsoft-com:vsta.v3");
                entryPointNode.SetAttribute("class", "PowerPointLabs.PostInstall");
                XmlElement postActionDataNode = doc.CreateElement("vstav3", "postActionData", "urn:schemas-microsoft-com:vsta.v3");
                
                addInNode.InsertAfter(postActionsNode, updateNode);
                postActionsNode.AppendChild(postActionNode);
                postActionNode.AppendChild(entryPointNode);
                entryPointNode.InnerXml = "<assemblyIdentity " +
                                          "name=" + AddQuote("PostInstall") + " " +
                                          "version=" + AddQuote(version) + " " +
                                          "language=" + AddQuote("neutral") + " " +
                                          "processorArchitecture=" + AddQuote("msil") + "/>";
                postActionNode.AppendChild(postActionDataNode);
                doc.Save(manifestDirectory);
            }
            else
            {
                DisplayWarning(ERROR_ALREADY_PATCHED);
                return;
            }

            //*************
            //***Re-Sign***
            //*************
            string argsForSignManifest =
                "-sign " + AddQuote(manifestDirectory) + " -certfile " + AddQuote(keyDirectory);
            string argsForSignVsto =
                "-update " + AddQuote(vstoDirectory) + " -appmanifest " + AddQuote(manifestDirectory) +
                " -certfile " + AddQuote(keyDirectory);
            System.Diagnostics.Process process = new Process();
            try
            {
                process.StartInfo.FileName = mageDirectory;
                process.StartInfo.Arguments = argsForSignManifest;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                process.WaitForExit();
                process = new Process();
                process.StartInfo.FileName = mageDirectory;
                process.StartInfo.Arguments = argsForSignVsto;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                process.WaitForExit();
            }
            catch
            {
                //Restore manifest file
                docBackup.Save(manifestDirectory);
                DisplayWarning(ERROR_INVALID_KEY_OR_MAGE_DIR);
                return;
            }
            System.IO.File.Copy(vstoDirectory, destVstoDirectory, true);

            DisplayDone();
        }
    }
}
