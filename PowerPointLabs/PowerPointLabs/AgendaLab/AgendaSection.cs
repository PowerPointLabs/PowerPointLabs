using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPointLabs.Utils;

namespace PowerPointLabs.AgendaLab
{
    internal struct AgendaSection
    {
        public readonly string Name;
        public readonly int Index;
        public readonly int Level;

        private readonly static char LevelDelimiter = '-';
        private readonly static char[] Delimiters = { LevelDelimiter };

        private readonly static int MaximumBulletIndent = 9;

        public AgendaSection(string name, int index, int level = 1)
        {
            Name = name;
            Index = index;
            Level = level;
        }

        public static AgendaSection FromSectionName(string name, int index)
        {
            return new AgendaSection(RemoveDelimiters(name), index, level: ParseSectionLevelFromName(name));
        }

        /// <summary>
        /// Note that a None section is encoded as a unique string of digits to prevent slide name collision
        /// </summary>
        public string Encode()
        {
            if (IsNone()) return Common.UniqueDigitString();
            return Name + "_" + Index;
        }

        public static AgendaSection Decode(string sectionStr)
        {
            int delimIndex = sectionStr.LastIndexOf("_");
            if (delimIndex == -1) return None;

            int index;
            bool result = int.TryParse(sectionStr.Substring(delimIndex + 1), out index);
            if (result == false) return None;

            string name = sectionStr.Substring(0, delimIndex);

            return new AgendaSection(name, index);
        }

        public bool IsNone()
        {
            return Name == null;
        }

        public static AgendaSection None
        {
            get { return new AgendaSection(); }
        }

        private static int ParseSectionLevelFromName(string section)
        {
            string newName = section.TrimStart(LevelDelimiter);
            int level = section.Length - newName.Length;
            return Math.Min(MaximumBulletIndent, level + 1);
        }

        private static string RemoveDelimiters(string originalName)
        {
            return originalName.TrimStart(Delimiters);
        }
    }
}
