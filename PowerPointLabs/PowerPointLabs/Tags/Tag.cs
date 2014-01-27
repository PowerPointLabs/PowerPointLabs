using System;
using System.Speech.Synthesis;

namespace AudioGen.Tags
{
    public abstract class Tag : ITag
    {
        public int Start { get; protected set; }
        public int End { get; protected set; }
        public string Contents { get; protected set; }
        public abstract bool Apply(PromptBuilder builder);
        public abstract string PrettyPrint();

        protected String ParseTagArgument()
        {
            int argumentStart = Contents.IndexOf(':') + 1;
            int argumentEnd = Contents.IndexOf(']');

            string argument = Contents.Substring(argumentStart, argumentEnd - argumentStart).Trim();
            return argument;
        }
    }
}
