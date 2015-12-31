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


            BindDataToPanels();

            var buttonHotkeyBindings = SetupButtonHotkeys();
            var hotkeyActionBindings = InitToolTipControl(buttonHotkeyBindings);
            InitialiseHotkeys(hotkeyActionBindings);
        }

        #region ToolTip
        private Dictionary<Native.VirtualKey, Action> InitToolTipControl(Dictionary<int, Native.VirtualKey> buttonHotkeyBindings)
        {
            var hotkeyActionBindings = new Dictionary<Native.VirtualKey, Action>();

            Action<ImageButton, Action, string> ConfigureButton = (button, action, tooltipMessage) =>
            {
                button.Click += (sender, e) => action();

                if (buttonHotkeyBindings.ContainsKey(button.ImageButtonUniqueId))
                {
                    var key = buttonHotkeyBindings[button.ImageButtonUniqueId];
                    tooltipMessage = "[" + VirtualKeyName(key) + "] " + tooltipMessage;
                    hotkeyActionBindings.Add(key, action);
                }

                SetTooltip(button, tooltipMessage);
            };


            // ---------------
            // || Top Panel ||
            // ---------------

            ConfigureButton(CircleButton, DrawingsLabMain.SwitchToCircleTool, "Draw a Circle.");
            ConfigureButton(TriangleButton, DrawingsLabMain.SwitchToTriangleTool, "Draw a Triangle.");

            ConfigureButton(RectButton, DrawingsLabMain.SwitchToRectangleTool, "Draw a Rectangle.");
            ConfigureButton(RoundedRectButton, DrawingsLabMain.SwitchToRoundedRectangleTool, "Draw a Rounded Rectangle.");

            ConfigureButton(LineButton, DrawingsLabMain.SwitchToLineTool, "Draw a Line.");
            ConfigureButton(ArrowButton, DrawingsLabMain.SwitchToArrowTool, "Draw a Curved Line.");

            ConfigureButton(TextboxButton, DrawingsLabMain.SwitchToTextboxTool, "Add a Text Box.");

            ConfigureButton(SelectTypeButton, DrawingsLabMain.SelectAllOfType, "Select all shapes of same type as currently selected shapes.");
            ConfigureButton(DuplicateButton, DrawingsLabMain.CloneTool, "Makes a copy of the selected shapes in the exact same location.");

            ConfigureButton(HideButton, DrawingsLabMain.HideTool, "Hide selected items.");
            ConfigureButton(ShowAllButton, DrawingsLabMain.ShowAllTool, "Show all hidden items.");
            ConfigureButton(SelectionPaneButton, DrawingsLabMain.OpenSelectionPane, "Opens the Selection Pane.");

            SetTooltip(ToggleHotkeysButton, "Enable / Disable Hotkeys.");


            // ---------------
            // || Tab: Main ||
            // ---------------

            ConfigureButton(AddTextButton, DrawingsLabMain.AddText, "Add Text to the selected shapes.");
            ConfigureButton(AddMathButton, DrawingsLabMain.AddMath, "Add Math to a selected shape.");
            ConfigureButton(RemoveTextButton, DrawingsLabMain.RemoveText, "Remove all text from the selected shapes.");

            ConfigureButton(GroupButton, DrawingsLabMain.GroupShapes, "Groups the selected shapes into a single shape.");
            ConfigureButton(UngroupButton, DrawingsLabMain.UngroupShapes, "Ungroups the selected group of shapes.");

            ConfigureButton(ArrowStartButton, DrawingsLabMain.ToggleArrowStart, "Groups the selected shapes into a single shape.");
            ConfigureButton(ArrowEndButton, DrawingsLabMain.ToggleArrowEnd, "Groups the selected shapes into a single shape.");

            ConfigureButton(MultiCloneExtendButton, DrawingsLabMain.MultiCloneExtendTool, "Extrapolates multiple copies of a shape, extending from two selected shapes.");
            ConfigureButton(MultiCloneBetweenButton, DrawingsLabMain.MultiCloneBetweenTool, "Interpolates multiple copies of a shape, in between two selected shapes.");
            ConfigureButton(MultiCloneGridButton, DrawingsLabMain.MultiCloneGridTool, "Extends two shapes into a grid of shapes.");
            
            ConfigureButton(BringForwardButton, DrawingsLabMain.BringForward, "Bring shapes Forward one step.");
            ConfigureButton(BringInFrontOfShapeButton, DrawingsLabMain.BringInFrontOfShape, "Bring shapes in front of last shape in selection.");
            ConfigureButton(BringToFrontButton, DrawingsLabMain.BringToFront, "Bring shapes to Front.");
            ConfigureButton(SendBackwardButton, DrawingsLabMain.SendBackward, "Send shapes Backward one step.");
            ConfigureButton(SendBehindShapeButton, DrawingsLabMain.SendBehindShape, "Send shapes behind last shape in selection.");
            ConfigureButton(SendToBackButton, DrawingsLabMain.SendToBack, "Send shapes to Back.");


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
            ConfigureButton(AlignBothButton, DrawingsLabMain.AlignBoth, "Align Shapes both Horizontally and Vertically to last shape in selection.");
            
            ConfigureButton(PivotAroundButton, DrawingsLabMain.PivotAroundTool, "Rotate / Multiclone a shape around another shape.");

            ConfigureButton(AlignHorizontalToSlideButton, DrawingsLabMain.AlignHorizontalToSlide, "Align Shapes Horizontally to a position relative to the slide.");
            ConfigureButton(AlignVerticalToSlideButton, DrawingsLabMain.AlignVerticalToSlide, "Align Shapes Vertically to a position relative to the slide.");
            ConfigureButton(AlignBothToSlideButton, DrawingsLabMain.AlignBothToSlide, "Align Shapes both Horizontally and Vertically to a position relative to the slide.");
            
            // --------------------
            // || Tab: Selection ||
            // --------------------



            return hotkeyActionBindings;
        }

        private void SetTooltip(DependencyObject button, string tooltipMessage)
        {
            ToolTip toolTip = new ToolTip { Content = tooltipMessage };
            ToolTipService.SetToolTip(button, toolTip);
        }

        private static string VirtualKeyName(Native.VirtualKey key)
        {
            var s = key.ToString();
            if (s.StartsWith("VK_")) s = s.Substring(3);
            return s;
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

        private Dictionary<int, Native.VirtualKey> SetupButtonHotkeys()
        {
            var bindings = new Dictionary<int, Native.VirtualKey>();
            Action<Native.VirtualKey, ImageButton> Assign = (key, button) =>
            {
                bindings.Add(button.ImageButtonUniqueId,key);
            };

            Assign(Native.VirtualKey.VK_L, LineButton);
            Assign(Native.VirtualKey.VK_R, RectButton);
            Assign(Native.VirtualKey.VK_C, CircleButton);
            Assign(Native.VirtualKey.VK_D, DuplicateButton);
            Assign(Native.VirtualKey.VK_A, SelectTypeButton);
            Assign(Native.VirtualKey.VK_H, HideButton);
            Assign(Native.VirtualKey.VK_F, BringForwardButton);
            Assign(Native.VirtualKey.VK_B, SendBackwardButton);

            return bindings;
        }

        private void InitialiseHotkeys(Dictionary<Native.VirtualKey, Action> hotkeyActionBindings)
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

            foreach (var entry in hotkeyActionBindings)
            {
                PPKeyboard.AddKeyupAction(entry.Key, RunOnlyWhenOpen(entry.Value));
            }
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

        private void TextColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog
            {
                Color = Graphics.ConvertRgbToColor(dataSource.FormatTextColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel) return;
            dataSource.FormatTextColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }
    }
}
