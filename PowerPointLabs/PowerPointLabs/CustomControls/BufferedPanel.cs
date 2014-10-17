using System.Windows.Forms;

namespace PowerPointLabs
{
    class BufferedPanel : Panel
    {
        public BufferedPanel()
        {
            DoubleBuffered = true;
            ResizeRedraw = true;
        }
    }
}
