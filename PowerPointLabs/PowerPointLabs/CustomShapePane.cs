using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Stepi.UI;

namespace PowerPointLabs
{
    public partial class CustomShapePane : UserControl
    {
        public CustomShapePane()
        {
            InitializeComponent();

            //foreach (var control in motherTableLayoutPanel.Controls)
            //{
            //    if (control is Panel)
            //    {
            //        var panel = control as Panel;
            //        var extendedPanel = panel.Controls[0] as ExtendedPanel;
            //        var rowIndex = motherTableLayoutPanel.GetRow(panel);

            //        motherTableLayoutPanel.RowStyles[rowIndex].Height = extendedPanel.Size.Height*
            //                                                            extendedPanel.CaptionSize/100;
            //    }
            //}
        }
    }
}
