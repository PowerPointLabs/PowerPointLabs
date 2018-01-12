using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace PowerPointLabs.XMLMisc
{
    class XmlParser
    {
        private Dictionary<string, string> shapeFileMapper;
        private Dictionary<string, string> audioIDFileMapper;

        private readonly XNamespace _p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private readonly XNamespace _r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private readonly XNamespace _a = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public XmlParser(string filePath)
        {
            shapeFileMapper = new Dictionary<string, string>();
            audioIDFileMapper = new Dictionary<string, string>();

            if (!File.Exists(filePath))
            {
                throw new ArgumentException("XML does not exist");
            }

            LinkShapeAndAudio(filePath);
        }

        public string GetCorrespondingAudio(string name)
        {
            return shapeFileMapper[name];
        }

        private void ParseRelation(string path)
        {
            var doc = File.ReadAllText(path);
            const string relaitonFormat = "<\\w+\\s\\w+=\\\"(\\w+\\d+)\\\" \\w+=\\\"[\\w\\:\\/\\.]+audio\\\" \\w+=\\\"[\\w\\.\\/]+(media\\d+\\.wav)\\\"\\/>";

            var regexRelation = new Regex(relaitonFormat);
            var matches = regexRelation.Matches(doc);

            for (int i = 0; i < matches.Count; i++)
            {
                var match = matches[i];

                audioIDFileMapper[match.Groups[1].Value] = match.Groups[2].Value;
            }
        }

        private void ParseShape(string path)
        {
            var doc = XDocument.Load(path);

            foreach (var element in doc.Descendants(_p + "spTree"))
            {
                var audioShape = element.Elements(_p + "pic");
                var pptSpeechFormat = new Regex("PowerPointLabs|AudioGen Speech \\d+");

                var data = from item in audioShape
                           where
                               pptSpeechFormat.IsMatch(item.Element(_p + "nvPicPr").Element(_p + "cNvPr").Attribute("name").Value)
                           select new
                                      {
                                          name = item.Element(_p + "nvPicPr").Element(_p + "cNvPr").Attribute("name").Value,
                                          audioID = item.Element(_p + "nvPicPr").Element(_p + "nvPr").Element(_a + "audioFile").Attribute(_r + "link").Value
                                      };

                foreach (var entry in data)
                {
                    shapeFileMapper[entry.name] = audioIDFileMapper[entry.audioID];
                }
            }
        }

        private void LinkShapeAndAudio(string path)
        {
            ParseRelation(path + ".rels");
            ParseShape(path);
        }
    }
}
