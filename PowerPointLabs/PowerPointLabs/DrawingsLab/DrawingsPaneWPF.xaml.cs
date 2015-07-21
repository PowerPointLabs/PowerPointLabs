using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using Shape = System.Windows.Shapes.Shape;

namespace PowerPointLabs.DrawingsLab
{
    /// <summary>
    /// Interaction logic for DrawingsPaneWPF.xaml
    /// </summary>
    public partial class DrawingsPaneWPF
    {
        public DrawingsPaneWPF()
        {
            InitializeComponent();
        }

        private void LineButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.SwitchToLineTool();
        }

        private void HideButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.HideTool();
        }

        private void CloneButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.CloneTool();
        }

        private void MultiCloneButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.MultiCloneTool();
        }
    }
}
