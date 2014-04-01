using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace DeployHelper
{
    class Program
    {
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

        static void Main(string[] args)
        {
            //TODO: refactor...
            //reference on what it does
            //http://msdn.microsoft.com/en-us/library/vstudio/dd465291(v=vs.100).aspx

            //read config
            Console.WriteLine("PowerPointLabs Version (Format - A.B.C.D, e.g. 1.7.1.0): ");
            string version = Console.ReadLine();
            while (string.IsNullOrEmpty(version))
            {
                version = Console.ReadLine();
            }
            string currentDirectory = System.Environment.CurrentDirectory;
            string configDirectory = currentDirectory + @"\DeployHelper.conf";
            string[] configContent;
            try
            {
                configContent = System.IO.File.ReadAllLines(configDirectory);
            }
            catch
            {
                DisplayWarning("Can't Find Config.");
                return;
            }
            string vstoDirectory = currentDirectory + @"\PowerPointLabs.vsto";
            string mageDirectory = configContent[1];
            string keyDirectory = configContent[3];
            string[] versionDetails = version.Split('.');
            string versionMajor;
            string versionMinor;
            string versionBuild;
            string versionRevision;
            try
            {
                versionMajor = versionDetails[0];
                versionMinor = versionDetails[1];
                versionBuild = versionDetails[2];
                versionRevision = versionDetails[3];
            }
            catch
            {
                DisplayWarning("Version Invalid.");
                return;
            }

            XmlDocument currentVsto;
            try
            {
                currentVsto = new XmlDocument();
                currentVsto.Load(vstoDirectory);
            }
            catch
            {
                DisplayWarning("Can't Find VSTO.");
                return;
            }
            var vstoNode = currentVsto.GetElementsByTagName("assemblyIdentity")[0];
            if (vstoNode.Attributes["version"].Value != version)
            {
                DisplayWarning("Different Version From Current VSTO's");
                return;
            }

            string buildDirectory = currentDirectory + @"\Application Files\PowerPointLabs_"
                + versionMajor + "_" + versionMinor + "_" + versionBuild + "_" + versionRevision;
            string manifestDirectory = buildDirectory + @"\PowerPointLabs.dll.manifest";
            string destVstoDirectory = buildDirectory + @"\PowerPointLabs.vsto";

            //modify manifest
            XmlDocument doc = new XmlDocument();
            XmlDocument docBackup = new XmlDocument();
            try
            {
                doc.Load(manifestDirectory);
                docBackup.Load(manifestDirectory);
            }
            catch
            {
                DisplayWarning("Can't Find Manifest For This Version");
                return;
            }
            //is patched?
            if (doc.GetElementsByTagName("vstav3:postAction").Count == 0)
            {
                //patch content
                string xmlFrag =
                    "<vstav3:postAction>" +
                    "<vstav3:entryPoint class=\"PowerPointLabs.PostInstall\">" +
                    "<assemblyIdentity name=\"PostInstall\" language=\"neutral\" " +
                    "version=\"" + version + "\" processorArchitecture=\"msil\" " +
                    "xmlns=\"urn:schemas-microsoft-com:asm.v2\" />" +
                    "</vstav3:entryPoint>" +
                    "<vstav3:postActionData>" +
                    "</vstav3:postActionData>" +
                    "</vstav3:postAction>";
                NameTable xmlFragNameTable = new NameTable();
                XmlNamespaceManager xmlFragNamespaceManager = new XmlNamespaceManager(xmlFragNameTable);
                xmlFragNamespaceManager.AddNamespace("vstav3", "urn:schemas-microsoft-com:vsta.v3");
                XmlParserContext xmlFragContext = new XmlParserContext(null, xmlFragNamespaceManager, null,
                    XmlSpace.None);
                XmlReaderSettings xmlFragSettings = new XmlReaderSettings();
                xmlFragSettings.ConformanceLevel = ConformanceLevel.Fragment;
                XmlReader xmlFragReader = XmlReader.Create(new StringReader(xmlFrag), xmlFragSettings, xmlFragContext);
                XmlDocument xmlFragDocument = new XmlDocument();
                xmlFragDocument.Load(xmlFragReader);
                XmlNode parentNode = doc.GetElementsByTagName("vstav3:addIn")[0];
                XmlNode node = doc.GetElementsByTagName("vstav3:update")[0];
                XmlElement nodeToInsert = doc.CreateElement("vstav3", "postActions", "urn:schemas-microsoft-com:vsta.v3");
                nodeToInsert.InnerXml = xmlFragDocument.InnerXml;
                parentNode.InsertAfter(nodeToInsert, node);
                doc.Save(manifestDirectory);
            }
            else
            {
                DisplayWarning("Already Patched.");
                return;
            }

            //re-sign
            string argumentsSignManifest =
                "-sign " + "\"" + manifestDirectory + "\"" + " -certfile " + "\"" + keyDirectory + "\"";
            string argumentsSignVsto =
                "-update " + "\"" + vstoDirectory + "\"" + " -appmanifest " + "\"" + manifestDirectory + "\"" +
                " -certfile " + "\"" + keyDirectory + "\"";
            System.Diagnostics.Process process = new Process();
            try
            {
                process.StartInfo.FileName = mageDirectory;
                process.StartInfo.Arguments = argumentsSignManifest;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                process.WaitForExit();
                process = new Process();
                process.StartInfo.FileName = mageDirectory;
                process.StartInfo.Arguments = argumentsSignVsto;
                process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                process.Start();
                process.WaitForExit();
            }
            catch
            {
                docBackup.Save(manifestDirectory);
                DisplayWarning("Invalid Mage or Key Directory.");
                return;
            }
            System.IO.File.Copy(vstoDirectory, destVstoDirectory, true);

            DisplayDone();
        }
    }
}
