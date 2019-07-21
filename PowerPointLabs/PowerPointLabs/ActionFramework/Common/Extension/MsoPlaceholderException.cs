using System;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    /// <summary>
    /// An exception that occurs due to a placeholder.
    /// </summary>
    class MsoPlaceholderException : Exception
    {
        public readonly PpPlaceholderType T;
        public MsoPlaceholderException(PpPlaceholderType t)
        {
            T = t;
        }

        public override string Message => T.ToString();
    }
}
