using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using PowerPointLabs.DataSources;
using PowerPointLabs.Models;
using PPExtraEventHelper;
using Shape = System.Windows.Shapes.Shape;

namespace PowerPointLabs.DrawingsLab
{
    /// <summary>
    /// Interaction logic for DrawingsPaneWPF.xaml
    /// </summary>
    public partial class DrawingsPaneWPF
    {
        private static bool hotkeysInitialised = false;

        public static DrawingsLabDataSource dataSource { get; private set; }

        public DrawingsPaneWPF()
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
            dataSource = FindResource("DrawingsLabData") as DrawingsLabDataSource;
            //ShiftValueX.DataContext = dataSource;
            //var binding = new Binding() {Path = new PropertyPath("ShiftValueX")};
            //this.ShiftValueX.SetBinding(ForegroundProperty, binding);
            //this.panel1.DataBindings.Add(new CustomBinding(
                //"BackColor",
                //dataSource,
                //"selectedColor",
                //new Converters.HSLColorToRGBColor()));

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

        // Block input when panel is open and user is not selecting text.
        private bool BlockInput()
        {
            return IsPanelOpen() &&
                   PowerPointCurrentPresentationInfo.CurrentSelection.Type != PpSelectionType.ppSelectionText;
        }

        private void InitialiseHotkeys()
        {
            if (hotkeysInitialised) return;
            hotkeysInitialised = true;

            PPKeyboard.AddConditionToBlockTextInput(BlockInput);

            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_L, RunOnlyWhenOpen(DrawingsLabMain.SwitchToLineTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_H, RunOnlyWhenOpen(DrawingsLabMain.HideTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_D, RunOnlyWhenOpen(DrawingsLabMain.CloneTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_F, RunOnlyWhenOpen(DrawingsLabMain.MultiCloneExtendTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_G, RunOnlyWhenOpen(DrawingsLabMain.MultiCloneBetweenTool));
        }
        #endregion

        private void LineButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.SwitchToLineTool();
        }

        private void HideButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.HideTool();
        }

        private void ShowAllButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.ShowAllTool();
        }

        private void CloneButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.CloneTool();
        }

        private void MultiCloneExtendButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.MultiCloneExtendTool();
        }

        private void MultiCloneBetweenButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.MultiCloneBetweenTool();
        }

        private void ApplyPositionButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.ApplyPosition();
        }

        private void RecordPositionButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.RecordPosition();
        }

        private void ApplyDisplacementButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.ApplyDisplacement();
        }

        private void RecordDisplacementButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.RecordDisplacement();
        }

        private void BringForwardButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.BringForward();
        }

        private void BringInFrontOfShapeButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.BringInFrontOfShape();
        }

        private void BringToFrontButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.BringToFront();
        }

        private void SendBackwardButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.SendBackward();
        }

        private void SendBehindShapeButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.SendBehindShape();
        }

        private void SendToBackButton_Click(object sender, EventArgs e)
        {
            DrawingsLabMain.SendToBack();
        }
    }
}
