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

        public static DrawingsLabDataSource DataSource { get; private set; }

        public DrawingsPaneWPF()
        {
            InitializeComponent();
            InitializeDataSource();

            var buttonHotkeyBindings = SetupButtonHotkeys();
            var hotkeyActionBindings = SetupButtons(buttonHotkeyBindings);
            InitializeHotkeys(hotkeyActionBindings);
        }

        #region ToolTip
        private Dictionary<Native.VirtualKey, Action> SetupButtons(Dictionary<int, Native.VirtualKey> buttonHotkeyBindings)
        {
            var hotkeyActionBindings = new Dictionary<Native.VirtualKey, Action>();

            Action<ImageButton, Action, string> configureButton = (button, action, tooltipMessage) =>
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

            configureButton(CircleButton, DrawingsLabMain.SwitchToCircleTool, "Draw a Circle.");
            configureButton(TriangleButton, DrawingsLabMain.SwitchToTriangleTool, "Draw a Triangle.");

            configureButton(RectButton, DrawingsLabMain.SwitchToRectangleTool, "Draw a Rectangle.");
            configureButton(RoundedRectButton, DrawingsLabMain.SwitchToRoundedRectangleTool, "Draw a Rounded Rectangle.");

            configureButton(LineButton, DrawingsLabMain.SwitchToLineTool, "Draw a Line.");
            configureButton(ArrowButton, DrawingsLabMain.SwitchToArrowTool, "Draw an Arrow.");

            configureButton(TextboxButton, DrawingsLabMain.SwitchToTextboxTool, "Add a Text Box.");

            configureButton(SelectTypeButton, DrawingsLabMain.SelectAllOfType, "Select all items of same type as the currently selected items.");
            configureButton(DuplicateButton, DrawingsLabMain.CloneTool, "Makes a copy of the selected items in the exact same location.");

            configureButton(HideButton, DrawingsLabMain.HideTool, "Hide selected items.");
            configureButton(ShowAllButton, DrawingsLabMain.ShowAllTool, "Show all hidden items.");
            configureButton(SelectionPaneButton, DrawingsLabMain.OpenSelectionPane, "Opens the Selection Pane.");

            SetTooltip(ToggleHotkeysButton, "Enable / Disable Hotkeys.");


            // ---------------
            // || Tab: Main ||
            // ---------------

            configureButton(AddTextButton, DrawingsLabMain.AddText, "Set the Text for the selected shapes.");
            configureButton(AddMathButton, DrawingsLabMain.AddMath, "Add Math to a selected shape, or convert highlighted text to Math.");
            configureButton(RemoveTextButton, DrawingsLabMain.RemoveText, "Remove all text from the selected shapes.");

            configureButton(GroupButton, DrawingsLabMain.GroupShapes, "Groups the selected shapes into a single shape.");
            configureButton(UngroupButton, DrawingsLabMain.UngroupShapes, "Ungroups the selected group of shapes.");

            configureButton(ArrowStartButton, DrawingsLabMain.ToggleArrowStart, "Toggles arrowheads at the start of the selected lines.");
            configureButton(ArrowEndButton, DrawingsLabMain.ToggleArrowEnd, "Toggles arrowheads at the end of the selected lines.");

            configureButton(MultiCloneExtendButton, DrawingsLabMain.MultiCloneExtendTool, "Extrapolates multiple copies of a shape, extending from two selected shapes.");
            configureButton(MultiCloneBetweenButton, DrawingsLabMain.MultiCloneBetweenTool, "Interpolates multiple copies of a shape, in between two selected shapes.");
            configureButton(MultiCloneGridButton, DrawingsLabMain.MultiCloneGridTool, "Extends two shapes into a grid of shapes.");
            
            configureButton(BringForwardButton, DrawingsLabMain.BringForward, "Bring shapes Forward one step.");
            configureButton(BringInFrontOfShapeButton, DrawingsLabMain.BringInFrontOfShape, "Bring shapes in front of last shape in the selection.");
            configureButton(BringToFrontButton, DrawingsLabMain.BringToFront, "Bring shapes to the Front.");
            configureButton(SendBackwardButton, DrawingsLabMain.SendBackward, "Send shapes Backward one step.");
            configureButton(SendBehindShapeButton, DrawingsLabMain.SendBehindShape, "Send shapes behind last shape in the selection.");
            configureButton(SendToBackButton, DrawingsLabMain.SendToBack, "Send shapes to the Back.");


            // -----------------
            // || Tab: Format ||
            // -----------------

            configureButton(ApplyFormatButton, ()=>DrawingsLabMain.ApplyFormat(applyAllSettings: false), "Apply recorded format to selected shapes.");
            configureButton(RecordFormatButton, DrawingsLabMain.RecordFormat, "Record Format of a selected shape.");
            

            // -------------------
            // || Tab: Position ||
            // -------------------

            configureButton(ApplyDisplacementButton, ()=>DrawingsLabMain.ApplyDisplacement(applyAllSettings: false), "Apply recorded displacement to selected shapes.");
            configureButton(ApplyPositionButton, () => DrawingsLabMain.ApplyPosition(applyAllSettings: false), "Apply recorded position or rotation to selected shapes.");
            configureButton(RecordDisplacementButton, DrawingsLabMain.RecordDisplacement, "Record the Displacement between two selected shapes.");
            configureButton(RecordPositionButton, DrawingsLabMain.RecordPosition, "Record position and rotation of a selected shape.");

            configureButton(AlignHorizontalButton, DrawingsLabMain.AlignHorizontal, "Align Shapes Horizontally to last shape in selection.");
            configureButton(AlignVerticalButton, DrawingsLabMain.AlignVertical, "Align Shapes Vertically to last shape in selection.");
            configureButton(AlignBothButton, DrawingsLabMain.AlignBoth, "Align Shapes both Horizontally and Vertically to last shape in selection.");
            
            configureButton(PivotAroundButton, DrawingsLabMain.PivotAroundTool, "Rotate / Multiclone a shape around another shape. Two shapes must be selected, the shape to be rotated and the pivot, in order.");

            configureButton(AlignHorizontalToSlideButton, DrawingsLabMain.AlignHorizontalToSlide, "Align Shapes Horizontally to a position relative to the slide.");
            configureButton(AlignVerticalToSlideButton, DrawingsLabMain.AlignVerticalToSlide, "Align Shapes Vertically to a position relative to the slide.");
            configureButton(AlignBothToSlideButton, DrawingsLabMain.AlignBothToSlide, "Align Shapes both Horizontally and Vertically to a position relative to the slide.");
            
            // --------------------
            // || Tab: Selection ||
            // --------------------

            // Empty as of now.

            return hotkeyActionBindings;
        }

        private void SetTooltip(DependencyObject button, string tooltipMessage)
        {
            ToolTip toolTip = new ToolTip { Content = tooltipMessage };
            ToolTipService.SetToolTip(button, toolTip);
        }

        /// <summary>
        /// Returns the string representaiton of the hotkey,
        /// to be used for hotkey displays in tooltips. E.g. "[L] Draw a Line"
        /// </summary>
        private static string VirtualKeyName(Native.VirtualKey key)
        {
            switch (key)
            {
                case Native.VirtualKey.VK_OEM_COMMA:
                    return " , ";
                case Native.VirtualKey.VK_OEM_PERIOD:
                    return " . ";
            }

            var s = key.ToString();
            if (s.StartsWith("VK_")) s = s.Substring(3);
            return s;
        }

        #endregion

        #region DataBindings
        private void InitializeDataSource()
        {
            DataSource = FindResource("DrawingsLabData") as DrawingsLabDataSource;
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
                   DataSource.HotkeysEnabled &&
                   PowerPointCurrentPresentationInfo.CurrentSelection.Type != PpSelectionType.ppSelectionText;
        }

        private Action RunOnlyWhenOpen(Action action)
        {
            return () => { if (IsReadingHotkeys()) action(); };
        }

        private Dictionary<int, Native.VirtualKey> SetupButtonHotkeys()
        {
            var usedKeys = new HashSet<Native.VirtualKey>();
            var bindings = new Dictionary<int, Native.VirtualKey>();
            Action<Native.VirtualKey, ImageButton> assign = (key, button) =>
            {
                bindings.Add(button.ImageButtonUniqueId, key);
                if (usedKeys.Contains(key)) throw new ArgumentException("Key already has a binding: " + key.ToString());
                usedKeys.Add(key);
            };

            assign(Native.VirtualKey.VK_L, LineButton);
            assign(Native.VirtualKey.VK_R, RectButton);
            assign(Native.VirtualKey.VK_E, RoundedRectButton);
            assign(Native.VirtualKey.VK_Y, TriangleButton);
            assign(Native.VirtualKey.VK_C, CircleButton);
            assign(Native.VirtualKey.VK_A, ArrowButton);
            assign(Native.VirtualKey.VK_T, TextboxButton);

            assign(Native.VirtualKey.VK_D, DuplicateButton);
            assign(Native.VirtualKey.VK_S, SelectTypeButton);
            assign(Native.VirtualKey.VK_H, HideButton);
            assign(Native.VirtualKey.VK_J, SelectionPaneButton);

            assign(Native.VirtualKey.VK_X, AddTextButton);
            assign(Native.VirtualKey.VK_G, GroupButton);
            assign(Native.VirtualKey.VK_U, UngroupButton);
            assign(Native.VirtualKey.VK_OEM_COMMA, ArrowStartButton);
            assign(Native.VirtualKey.VK_OEM_PERIOD, ArrowEndButton);

            assign(Native.VirtualKey.VK_N, MultiCloneBetweenButton);
            assign(Native.VirtualKey.VK_M, MultiCloneExtendButton);

            assign(Native.VirtualKey.VK_O, AlignBothButton);
            assign(Native.VirtualKey.VK_P, PivotAroundButton);

            assign(Native.VirtualKey.VK_F, BringForwardButton);
            assign(Native.VirtualKey.VK_B, SendBackwardButton);

            return bindings;
        }

        private void InitializeHotkeys(Dictionary<Native.VirtualKey, Action> hotkeyActionBindings)
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
                Color = Graphics.ConvertRgbToColor(DataSource.FormatFillColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel) return;
            DataSource.FormatFillColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }

        private void LineColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog
            {
                Color = Graphics.ConvertRgbToColor(DataSource.FormatLineColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel) return;
            DataSource.FormatLineColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }

        private void TextColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog
            {
                Color = Graphics.ConvertRgbToColor(DataSource.FormatTextColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel) return;
            DataSource.FormatTextColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }
    }
}
