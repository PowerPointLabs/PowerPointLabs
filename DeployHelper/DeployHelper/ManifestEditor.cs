using System;
using System.IO;
using System.Xml;

namespace DeployHelper
{
    class ManifestEditor
    {
        #region Modify Manifest

        private XmlDocument _manifest = new XmlDocument();
        private XmlDocument _manifestBackup = new XmlDocument();
        private string _manifestDirectory;

        public ManifestEditor(String manifestDirectory)
        {
            Console.WriteLine("Patching...");
            _manifestDirectory = manifestDirectory;

            VerifyManifestExist();
            _manifest.Load(manifestDirectory);
            _manifestBackup.Load(manifestDirectory);
        }

        private void VerifyManifestExist()
        {
            if (!File.Exists(_manifestDirectory))
            {
                throw new FileNotFoundException();
            }
        }

        public Boolean IsPatched()
        {
            return _manifest.GetElementsByTagName("vstav3:postAction").Count != 0;
        }

        public void RestoreManifest()
        {
            _manifestBackup.Save(_manifestDirectory);
        }

        public void ModifyManifest()
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
            const string vstaNamespaceUri = "urn:schemas-microsoft-com:vsta.v3";

            //setup nodes
            XmlNode addInNode = _manifest.GetElementsByTagName("vstav3:addIn")[0];
            XmlNode updateNode = _manifest.GetElementsByTagName("vstav3:update")[0];
            XmlElement postActionsNode = _manifest.CreateElement("vstav3", "postActions",
                vstaNamespaceUri);
            XmlElement postActionNode = _manifest.CreateElement("vstav3", "postAction",
                vstaNamespaceUri);
            XmlElement entryPointNode = _manifest.CreateElement("vstav3", "entryPoint",
                vstaNamespaceUri);
            entryPointNode.SetAttribute("class", "PowerPointLabs.PostInstall");
            XmlElement postActionDataNode = _manifest.CreateElement("vstav3", "postActionData",
                vstaNamespaceUri);

            //insert and append those nodes
            addInNode.InsertAfter(postActionsNode, updateNode);
            postActionsNode.AppendChild(postActionNode);
            postActionNode.AppendChild(entryPointNode);
            entryPointNode.InnerXml = "<assemblyIdentity " +
                                      "name=" + Util.AddQuote("PostInstall") + " " +
                                      "version=" + Util.AddQuote("1.0.0.0") + " " +
                                      "language=" + Util.AddQuote("neutral") + " " +
                                      "processorArchitecture=" + Util.AddQuote("msil") + "/>";
            postActionNode.AppendChild(postActionDataNode);
            _manifest.Save(_manifestDirectory);
        }
        #endregion
    }
}
