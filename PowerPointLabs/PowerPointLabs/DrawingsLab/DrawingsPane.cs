using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ColorPicker;
using PowerPointLabs.DataSources;
using PowerPointLabs.DrawingsLab;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using PPExtraEventHelper;

using Converters = PowerPointLabs.Converters;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{

    public partial class DrawingsPane : UserControl
    {
        public DrawingsPane()
        {
            InitializeComponent();
        }
    }
}
