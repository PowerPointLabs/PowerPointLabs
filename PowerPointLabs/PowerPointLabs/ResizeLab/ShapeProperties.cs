using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.ResizeLab
{
    public class ShapeProperties
    {
        public string Name { get; }
        public float Top { get; }
        public float Left { get; }
        public float AbsoluteWidth { get; }
        public float AbsoluteHeight { get; }

        public ShapeProperties(string name, float top, float left, float absoluteWidth, float absoluteHeight)
        {
            Name = name;
            Top = top;
            Left = left;
            AbsoluteWidth = absoluteWidth;
            AbsoluteHeight = absoluteHeight;
        }
    }
}
