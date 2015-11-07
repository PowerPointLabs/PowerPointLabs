using System;

namespace PowerPointLabs.ImageSearch.Domain
{
    [Serializable]
    public class WindowInfo
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        public WindowInfo()
        {
            Left = -1;
            Top = -1;
            Width = -1;
            Height = -1;
        }
    }
}
