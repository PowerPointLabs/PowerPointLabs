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

        public AgendaSection(string name, int index)
        {
            Name = name;
            Index = index;
        }

        /// <summary>
        /// Note that a None section is encoded as a unique string of digits to prevent slide name collision
        /// </summary>
        public string Encode()
        {
            if (IsNone())
            {
                return Common.UniqueDigitString();
            }
            return Name + "_" + Index;
        }

        public static AgendaSection Decode(string sectionStr)
        {
            int delimIndex = sectionStr.LastIndexOf("_");
            if (delimIndex == -1)
            {
                return None;
            }

            int index;
            bool result = int.TryParse(sectionStr.Substring(delimIndex + 1), out index);
            if (result == false)
            {
                return None;
            }

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
    }
}
