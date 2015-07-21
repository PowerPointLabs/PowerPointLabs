using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PowerPointLabs.DataSources;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using System.Drawing.Drawing2D;
using PPExtraEventHelper;
using Converters = PowerPointLabs.Converters;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ColorPicker;
using PowerPointLabs.Views;
using Microsoft.Office.Core;
using PowerPointLabs.DrawingsLab;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{

    public partial class DrawingsPane : UserControl
    {
        private static bool hotkeysInitialised = false;

        public DrawingsPane()
        {
            InitializeComponent();

            InitialiseHotkeys();

            BindDataToPanels();

            InitToolTipControl();
        }

        #region ToolTip
        private void InitToolTipControl()
        {
            //toolTip1.SetToolTip(panel1, TextCollection.ColorsLabText.MainColorBoxTooltips);
        }
        #endregion

        #region DataBindings
        private void BindDataToPanels()
        {
            //this.panel1.DataBindings.Add(new CustomBinding(
                //"BackColor",
                //dataSource,
                //"selectedColor",
                //new Converters.HSLColorToRGBColor()));
        }
        #endregion

        #region ButtonCallbacks
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
        #endregion

        #region HotkeyInitialisation
        private bool IsPanelOpen()
        {
            var drawingsPane = Globals.ThisAddIn.GetActivePane(typeof(DrawingsPane));
            return drawingsPane.Visible;
        }

        private Action RunOnlyWhenOpen(Action action)
        {
            return () => { if (IsPanelOpen()) action(); };
        }

        private void InitialiseHotkeys()
        {
            if (hotkeysInitialised) return;
            hotkeysInitialised = true;

            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_L, RunOnlyWhenOpen(DrawingsLabMain.SwitchToLineTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_H, RunOnlyWhenOpen(DrawingsLabMain.HideTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_D, RunOnlyWhenOpen(DrawingsLabMain.CloneTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_F, RunOnlyWhenOpen(DrawingsLabMain.MultiCloneTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_Z, RunOnlyWhenOpen(() =>
            {
                var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
                if (selection.Type != PpSelectionType.ppSelectionShapes) return;

                var shapes = selection.ShapeRange;
                if (shapes.Count != 1)
                {
                    return;
                }
                var shape = shapes[1];

                Debug.WriteLine(shape.Width + " , " + shape.Height);
            }
                ));
        }
        #endregion


        protected override CreateParams CreateParams
        {
            get
            {
                var createParams = base.CreateParams;
                createParams.ExStyle |= (int)Native.Message.WS_EX_COMPOSITED;  // Turn on WS_EX_COMPOSITED
                return createParams;
            }
        }
    }
}
