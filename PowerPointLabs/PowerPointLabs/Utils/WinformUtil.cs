using System.Drawing;

namespace PowerPointLabs.Utils
{
    public class WinformUtil
    {
        public static Point MousePosition => System.Windows.Forms.Control.MousePosition;
        public static Size WorkingAreaSize => System.Windows.Forms.SystemInformation.WorkingArea.Size;
    }
}
