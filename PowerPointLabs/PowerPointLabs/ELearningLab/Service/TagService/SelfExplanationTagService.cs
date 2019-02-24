using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Service
{
    public class SelfExplanationTagService
    {
        private static HashSet<int> tags = new HashSet<int>();

        public static void PopulateTagNos(List<string> shapeNames)
        {
            foreach (string shapeName in shapeNames)
            {
                ExtractTagNo(shapeName);
            }
        }


        public static int GenerateUniqueTag()
        {
            int count = 0;
            do
            {
                count++;
            }
            while (tags.Contains(count));

            tags.Add(count);
            return count;
        }

        public static int ExtractTagNo(string name)
        {
            Regex regex = new Regex(ELearningLabText.ExtractTagNoRegex, RegexOptions.IgnoreCase);
            Match match = regex.Match(name);
            int value = -1;
            if (match.Success)
            {
                try
                {
                    value = int.Parse(match.Groups[1].Value.Trim());
                    tags.Add(value);
                }
                catch (Exception e)
                {
                    Logger.Log(e.Message);
                }
            }
            return value;
        }

        public static void Clear()
        {
            tags.Clear();
        }
    }
}
