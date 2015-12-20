using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.DataSources;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.WPF;
using PPExtraEventHelper;
using Shape = System.Windows.Shapes.Shape;
using ToolTip = System.Windows.Controls.ToolTip;

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
            // ---------------
            // || Top Panel ||
            // ---------------

            ConfigureButton(CircleButton, DrawingsLabMain.SwitchToCircleTool, "[C] Draw a Circle.");
            ConfigureButton(RectButton, DrawingsLabMain.SwitchToRectangleTool, "[R] Draw a Rectangle.");
            ConfigureButton(RoundedRectButton, DrawingsLabMain.SwitchToRoundedRectangleTool, "Draw a Rounded Rectangle.");

            ConfigureButton(LineButton, DrawingsLabMain.SwitchToLineTool, "[L] Draw a Line.");
            ConfigureButton(CurveButton, DrawingsLabMain.SwitchToCurvedLineTool, "Draw a Curved Line.");
            // SetTooltip ToggleArrowsButton

            // ConfigureButton TextboxButton
            // ConfigureButton MathboxButton

            ConfigureButton(SelectTypeButton, DrawingsLabMain.SelectAllOfType, "[A] Select all shapes of same type as currently selected shapes.");
            ConfigureButton(HideButton, DrawingsLabMain.HideTool, "[H] Hide selected items.");
            ConfigureButton(ShowAllButton, DrawingsLabMain.ShowAllTool, "[S] Show all hidden items.");
            ConfigureButton(DuplicateButton, DrawingsLabMain.CloneTool, "[D] Makes a copy of the selected shapes in the exact same location.");
            

            SetTooltip(ToggleHotkeysButton, "Enable / Disable Hotkeys.");


            // ---------------
            // || Tab: Main ||
            // ---------------

            ConfigureButton(AddTextButtonMain, DrawingsLabMain.AddText, "Add Text to the selected shapes.");
            ConfigureButton(AddMathButtonMain, DrawingsLabMain.AddMath, "Add Math");
            // ConfigureButton AddMathButtonMain
            // ConfigureButton RemoveTextButtonMain

            // ConfigureButton GroupButtonMain
            // ConfigureButton UngroupButtonMain

            ConfigureButton(ApplyDisplacementButtonMain, ()=>DrawingsLabMain.ApplyDisplacement(applyAllSettings: true), "Apply recorded displacement to selected shapes.");
            ConfigureButton(ApplyFormatButtonMain, ()=>DrawingsLabMain.ApplyFormat(applyAllSettings: true), "Apply recorded format to selected shapes.");
            ConfigureButton(ApplyPositionButtonMain, ()=>DrawingsLabMain.ApplyPosition(applyAllSettings: true), "Apply recorded position or rotation to selected shapes.");
            ConfigureButton(RecordDisplacementButtonMain, DrawingsLabMain.RecordDisplacement, "Record Displacement between two selected shapes.");
            ConfigureButton(RecordFormatButtonMain, DrawingsLabMain.RecordFormat, "Record Format of a selected shape.");
            ConfigureButton(RecordPositionButtonMain, DrawingsLabMain.RecordPosition, "Record position and rotation of a selected shape.");

            ConfigureButton(AlignHorizontalButtonMain, DrawingsLabMain.AlignHorizontal, "Align Shapes Horizontally to last shape in selection.");
            ConfigureButton(AlignVerticalButtonMain, DrawingsLabMain.AlignVertical, "Align Shapes Vertically to last shape in selection.");

            ConfigureButton(MultiCloneExtendButtonMain, DrawingsLabMain.MultiCloneExtendTool, "[N] Extrapolates multiple copies of a shape, extending from two selected shapes.");
            ConfigureButton(MultiCloneBetweenButtonMain, DrawingsLabMain.MultiCloneBetweenTool, "[M] Interpolates multiple copies of a shape, in between two selected shapes.");

            ConfigureButton(BringForwardButtonMain, DrawingsLabMain.BringForward, "[F] Bring shapes Forward one step.");
            ConfigureButton(BringInFrontOfShapeButtonMain, DrawingsLabMain.BringInFrontOfShape, "Bring shapes in front of last shape in selection.");
            ConfigureButton(BringToFrontButtonMain, DrawingsLabMain.BringToFront, "Bring shapes to Front.");
            ConfigureButton(SendBackwardButtonMain, DrawingsLabMain.SendBackward, "[B] Send shapes Backward one step.");
            ConfigureButton(SendBehindShapeButtonMain, DrawingsLabMain.SendBehindShape, "Send shapes behind last shape in selection.");
            ConfigureButton(SendToBackButtonMain, DrawingsLabMain.SendToBack, "Send shapes to Back.");


            // -----------------
            // || Tab: Format ||
            // -----------------

            ConfigureButton(ApplyFormatButton, ()=>DrawingsLabMain.ApplyFormat(applyAllSettings:false), "Apply recorded format to selected shapes.");
            ConfigureButton(RecordFormatButton, DrawingsLabMain.RecordFormat, "Record Format of a selected shape.");
            

            // -------------------
            // || Tab: Position ||
            // -------------------

            ConfigureButton(ApplyDisplacementButton, ()=>DrawingsLabMain.ApplyDisplacement(applyAllSettings:false), "Apply recorded displacement to selected shapes.");
            ConfigureButton(ApplyPositionButton, () => DrawingsLabMain.ApplyPosition(applyAllSettings: false), "Apply recorded position or rotation to selected shapes.");
            ConfigureButton(RecordDisplacementButton, DrawingsLabMain.RecordDisplacement, "Record Displacement between two selected shapes.");
            ConfigureButton(RecordPositionButton, DrawingsLabMain.RecordPosition, "Record position and rotation of a selected shape.");

            ConfigureButton(AlignHorizontalButton, DrawingsLabMain.AlignHorizontal, "Align Shapes Horizontally to last shape in selection.");
            ConfigureButton(AlignVerticalButton, DrawingsLabMain.AlignVertical, "Align Shapes Vertically to last shape in selection.");
            // ConfigureButton AlignBothButton

            ConfigureButton(AlignHorizontalToSlideButton, DrawingsLabMain.AlignHorizontalToSlide, "Align Shapes Horizontally to a position relative to the slide.");
            ConfigureButton(AlignVerticalToSlideButton, DrawingsLabMain.AlignVerticalToSlide, "Align Shapes Vertically to a position relative to the slide.");
            // ConfigureButton AlignBothToSlideButton

            ConfigureButton(MultiCloneExtendButton, DrawingsLabMain.MultiCloneExtendTool, "[N] Extrapolates multiple copies of a shape, extending from two selected shapes.");
            ConfigureButton(MultiCloneBetweenButton, DrawingsLabMain.MultiCloneBetweenTool, "[M] Interpolates multiple copies of a shape, in between two selected shapes.");
            // ConfigureButton MultiCloneGridButton
            // ConfigureButton PivotAroundButton

            ConfigureButton(BringForwardButton, DrawingsLabMain.BringForward, "[F] Bring shapes Forward one step.");
            ConfigureButton(BringInFrontOfShapeButton, DrawingsLabMain.BringInFrontOfShape, "Bring shapes in front of last shape in selection.");
            ConfigureButton(BringToFrontButton, DrawingsLabMain.BringToFront, "Bring shapes to Front.");
            ConfigureButton(SendBackwardButton, DrawingsLabMain.SendBackward, "[B] Send shapes Backward one step.");
            ConfigureButton(SendBehindShapeButton, DrawingsLabMain.SendBehindShape, "Send shapes behind last shape in selection.");
            ConfigureButton(SendToBackButton, DrawingsLabMain.SendToBack, "Send shapes to Back.");


            // --------------------
            // || Tab: Selection ||
            // --------------------

        }

        private void ConfigureButton(ImageButton button, Action action, string tooltipMessage)
        {
            button.Click += (sender, e) => action();
            
            ToolTip toolTip = new ToolTip { Content = tooltipMessage };
            ToolTipService.SetToolTip(button, toolTip);
        }

        private void SetTooltip(DependencyObject button, string tooltipMessage)
        {
            ToolTip toolTip = new ToolTip { Content = tooltipMessage };
            ToolTipService.SetToolTip(button, toolTip);
        }

        #endregion

        #region DataBindings
        private void BindDataToPanels()
        {
            dataSource = FindResource("DrawingsLabData") as DrawingsLabDataSource;
        }
        #endregion


        #region HotkeyInitialisation
        private bool IsPanelOpen()
        {
            var drawingsPane = Globals.ThisAddIn.GetActivePane(typeof(DrawingsPane));
            return drawingsPane.Visible;
        }

        private bool IsReadingHotkeys()
        {
            // Is reading hotkeys when panel is open and user is not selecting text.
            return IsPanelOpen() &&
                   dataSource.HotkeysEnabled &&
                   PowerPointCurrentPresentationInfo.CurrentSelection.Type != PpSelectionType.ppSelectionText;
        }

        private Action RunOnlyWhenOpen(Action action)
        {
            return () => { if (IsReadingHotkeys()) action(); };
        }

        private void InitialiseHotkeys()
        {
            if (hotkeysInitialised) return;
            hotkeysInitialised = true;

            PPKeyboard.AddConditionToBlockTextInput(IsReadingHotkeys);

            var numericKeys = new[]
            {
                Native.VirtualKey.VK_0,
                Native.VirtualKey.VK_1,
                Native.VirtualKey.VK_2,
                Native.VirtualKey.VK_3,
                Native.VirtualKey.VK_4,
                Native.VirtualKey.VK_5,
                Native.VirtualKey.VK_6,
                Native.VirtualKey.VK_7,
                Native.VirtualKey.VK_8,
                Native.VirtualKey.VK_9,
            };

            // I use a regular for loop due to inconsistent compiler behaviour when foreach is used.
            for (int i = 0; i < numericKeys.Length; ++i)
            {
                var key = numericKeys[i];
                // Assign number and ctrl+number to control group commands.
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => DrawingsLabMain.SelectControlGroup(key)));
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => DrawingsLabMain.SetControlGroup(key)), ctrl: true);

                // Block shift+number and ctrl+shift+number
                PPKeyboard.AddConditionToBlockTextInput(IsReadingHotkeys, key, shift: true);
                PPKeyboard.AddConditionToBlockTextInput(IsReadingHotkeys, key, ctrl: true, shift: true);

                // Assign shift+number and ctrl+shift+number to control group commands.
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => DrawingsLabMain.SelectControlGroup(key, appendToSelection: true)), shift: true);
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => DrawingsLabMain.SetControlGroup(key, appendToGroup: true)), ctrl: true, shift: true);
            }

            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_L, RunOnlyWhenOpen(DrawingsLabMain.SwitchToLineTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_R, RunOnlyWhenOpen(DrawingsLabMain.SwitchToRectangleTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_C, RunOnlyWhenOpen(DrawingsLabMain.SwitchToCircleTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_D, RunOnlyWhenOpen(DrawingsLabMain.CloneTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_A, RunOnlyWhenOpen(DrawingsLabMain.SelectAllOfType));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_H, RunOnlyWhenOpen(DrawingsLabMain.HideTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_S, RunOnlyWhenOpen(DrawingsLabMain.ShowAllTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_M, RunOnlyWhenOpen(DrawingsLabMain.MultiCloneExtendTool));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_N, RunOnlyWhenOpen(DrawingsLabMain.MultiCloneBetweenTool));

            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_F, RunOnlyWhenOpen(DrawingsLabMain.BringForward));
            PPKeyboard.AddKeyupAction(Native.VirtualKey.VK_B, RunOnlyWhenOpen(DrawingsLabMain.SendBackward));
        }
        #endregion

        private void FillColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog
            {
                Color = Graphics.ConvertRgbToColor(dataSource.FormatFillColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel) return;
            dataSource.FormatFillColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }

        private void LineColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog
            {
                Color = Graphics.ConvertRgbToColor(dataSource.FormatLineColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel) return;
            dataSource.FormatLineColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }
    }
}
