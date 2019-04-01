using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Service
{
    public class ELearningLabTextStorageService
    {
        public static void StoreSelfExplanationTextToSlide(List<ExplanationItem> selfExplanationClickItems, 
            PowerPointSlide slide)
        {
            string shapeName = ELearningLabText.ELearningLabTextStorageShapeName;
            slide.DeleteShapeWithName(shapeName);
            List<Dictionary<string, string>> selfExplanationText =
                ConvertListToDictionary(selfExplanationClickItems);
            XElement textInxml = new XElement(ELearningLabText.SelfExplanationTextIdentifier, 
                selfExplanationText.Select( kv =>
                 new XElement(ELearningLabText.SelfExplanationItemIdentifier, 
                from text in kv select new XElement(text.Key, text.Value))));
            Shape shape = ShapeUtility.InsertSelfExplanationTextBoxToSlide(slide, shapeName, textInxml.ToString());
        }

        public static List<Dictionary<string, string>> LoadSelfExplanationsFromSlide(PowerPointSlide slide)
        {
            List<Shape> shapes = slide.GetShapeWithName(ELearningLabText.ELearningLabTextStorageShapeName);
            if (shapes.Count > 0)
            {
                Shape shape = shapes[0];
                return LoadSelfExplanationTextFromString(shape.TextFrame.TextRange.Text);
            }
            return null;
        }

        private static List<Dictionary<string, string>> LoadSelfExplanationTextFromString(string text)
        {
            List<Dictionary<string, string>> tagNoToSelfExplanationTextDic = 
                new List<Dictionary<string, string>>();
            XElement xml = XElement.Parse(text);
            foreach (var selfExplanation in xml.Elements())
            {
                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic.Add(ELearningLabText.CaptionTextIdentifier, 
                    selfExplanation.Element(ELearningLabText.CaptionTextIdentifier).Value);
                string calloutText = selfExplanation.Element(ELearningLabText.CalloutTextIdentifier).Value;
                if (string.IsNullOrEmpty(calloutText.Trim()))
                {
                    dic.Add(ELearningLabText.CalloutTextIdentifier, 
                        selfExplanation.Element(ELearningLabText.CaptionTextIdentifier).Value);
                }
                else
                {
                    dic.Add(ELearningLabText.CalloutTextIdentifier, 
                        selfExplanation.Element(ELearningLabText.CalloutTextIdentifier).Value);
                }
                dic.Add(ELearningLabText.TagNoIdentifier, 
                    selfExplanation.Element(ELearningLabText.TagNoIdentifier).Value);
                tagNoToSelfExplanationTextDic.Add(dic);
            }
            return tagNoToSelfExplanationTextDic;
        }

        private static List<Dictionary<string, string>> ConvertListToDictionary(List<ExplanationItem> selfExplanationClickItems)
        {
            List<Dictionary<string, string>> keyValuePairs = 
                new List<Dictionary<string, string>>();
            foreach (ExplanationItem item in selfExplanationClickItems)
            {
                Dictionary<string, string> value = new Dictionary<string, string>();
                value.Add(ELearningLabText.CaptionTextIdentifier, item.CaptionText);
                if (!item.CaptionText.Trim().Equals(item.CalloutText.Trim()))
                {
                    value.Add(ELearningLabText.CalloutTextIdentifier, item.CalloutText);
                }
                else
                {
                    value.Add(ELearningLabText.CalloutTextIdentifier, string.Empty);
                }
                value.Add(ELearningLabText.TagNoIdentifier, item.tagNo.ToString());
                keyValuePairs.Add(value);
            }
            return keyValuePairs;
        }
    }
}
