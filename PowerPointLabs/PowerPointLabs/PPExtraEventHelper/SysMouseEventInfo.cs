using System;

namespace PPExtraEventHelper
{
    [Obsolete("DO NOT use this class! Instead, use PPMouse.")]
    public class SysMouseEventInfo : EventArgs
    {
        public string WindowTitle { get; set; }
    }
}
