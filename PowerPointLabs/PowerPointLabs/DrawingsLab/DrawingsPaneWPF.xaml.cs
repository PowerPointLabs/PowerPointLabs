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
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.DataSources;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.WPF;
using PPExtraEventHelper;
using MessageBox = System.Windows.Forms.MessageBox;
using Shape = System.Windows.Shapes.Shape;
using ToolTip = System.Windows.Controls.ToolTip;

namespace PowerPointLabs.DrawingsLab
{
    /// <summary>
    /// Interaction logic for DrawingsPaneWPF.xaml
    /// </summary>
    public partial class DrawingsPaneWPF
    {
        private DrawingLabData _data;
        private DrawingsLabMain _drawingLab;
        private readonly DrawingsLabDataSource _dataSource;

        public DrawingsPaneWPF()
        {
            InitializeComponent();
            _dataSource = FindResource("DrawingsLabData") as DrawingsLabDataSource;
        }

        #region Data Binding
        internal void TryInitialise(DrawingLabData data, DrawingsLabMain drawingLab)
        {
            if (_data != null)
            {
                return;
            }

            _data = data;
            _dataSource.AssignData(data);
            _drawingLab = drawingLab;

            InitialiseButtonsAndHotkeys();
        }

        private void InitialiseButtonsAndHotkeys()
        {
            var buttonHotkeyBindings = SetupButtonHotkeys();
            var hotkeyActionBindings = SetupButtons(buttonHotkeyBindings);
            InitializeHotkeys(hotkeyActionBindings);
        }

        #endregion


        #region Button Bindings / ToolTips
        private Dictionary<Native.VirtualKey, Action> SetupButtons(Dictionary<int, Native.VirtualKey> buttonHotkeyBindings)
        {
            var hotkeyActionBindings = new Dictionary<Native.VirtualKey, Action>();

            Action<ImageButton, Action, string> configureButton = (button, action, tooltipMessage) =>
            {
                action = _drawingLab.FunctionWrapper(action);
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

            configureButton(CircleButton, _drawingLab.SwitchToCircleTool, "Draw a Circle.");
            configureButton(TriangleButton, _drawingLab.SwitchToTriangleTool, "Draw a Triangle.");

            configureButton(RectButton, _drawingLab.SwitchToRectangleTool, "Draw a Rectangle.");
            configureButton(RoundedRectButton, _drawingLab.SwitchToRoundedRectangleTool, "Draw a Rounded Rectangle.");

            configureButton(LineButton, _drawingLab.SwitchToLineTool, "Draw a Line.");
            configureButton(ArrowButton, _drawingLab.SwitchToArrowTool, "Draw an Arrow.");

            configureButton(TextboxButton, _drawingLab.SwitchToTextboxTool, "Add a Text Box.");

            configureButton(SelectTypeButton, _drawingLab.SelectAllOfType, "Select all items of same type as the currently selected items.");
            configureButton(DuplicateButton, _drawingLab.CloneTool, "Makes a copy of the selected items in the exact same location.");

            configureButton(HideButton, _drawingLab.HideTool, "Hide selected items.");
            configureButton(ShowAllButton, _drawingLab.ShowAllTool, "Show all hidden items.");
            configureButton(SelectionPaneButton, _drawingLab.OpenSelectionPane, "Opens the Selection Pane.");

            SetTooltip(ToggleHotkeysButton, "Enable / Disable Hotkeys.");


            // ---------------
            // || Tab: Main ||
            // ---------------

            configureButton(AddTextButton, _drawingLab.AddText, "Set the Text for the selected shapes.");
            configureButton(AddMathButton, _drawingLab.AddMath, "Add Math to a selected shape, or convert highlighted text to Math.");
            configureButton(RemoveTextButton, _drawingLab.RemoveText, "Remove all text from the selected shapes.");

            configureButton(GroupButton, _drawingLab.GroupShapes, "Groups the selected shapes into a single shape.");
            configureButton(UngroupButton, _drawingLab.UngroupShapes, "Ungroups the selected group of shapes.");

            configureButton(ArrowStartButton, _drawingLab.ToggleArrowStart, "Toggles arrowheads at the start of the selected lines.");
            configureButton(ArrowEndButton, _drawingLab.ToggleArrowEnd, "Toggles arrowheads at the end of the selected lines.");

            configureButton(MultiCloneExtendButton, _drawingLab.MultiCloneExtendTool, "Extrapolates multiple copies of a shape, extending from two selected shapes.");
            configureButton(MultiCloneBetweenButton, _drawingLab.MultiCloneBetweenTool, "Interpolates multiple copies of a shape, in between two selected shapes.");
            configureButton(MultiCloneGridButton, _drawingLab.MultiCloneGridTool, "Extends two shapes into a grid of shapes.");
            
            configureButton(BringForwardButton, _drawingLab.BringForward, "Bring shapes Forward one step.");
            configureButton(BringInFrontOfShapeButton, _drawingLab.BringInFrontOfShape, "Bring shapes in front of last shape in the selection.");
            configureButton(BringToFrontButton, _drawingLab.BringToFront, "Bring shapes to the Front.");
            configureButton(SendBackwardButton, _drawingLab.SendBackward, "Send shapes Backward one step.");
            configureButton(SendBehindShapeButton, _drawingLab.SendBehindShape, "Send shapes behind last shape in the selection.");
            configureButton(SendToBackButton, _drawingLab.SendToBack, "Send shapes to the Back.");


            // -----------------
            // || Tab: Format ||
            // -----------------

            configureButton(ApplyFormatButton, () => _drawingLab.ApplyFormat(applyAllSettings: false), "Apply recorded format to selected shapes.");
            configureButton(RecordFormatButton, _drawingLab.RecordFormat, "Record Format of a selected shape.");
            

            Func<string, string> toggleSyncButtonToolTip = type => "Enable if you want to sync the " + type + " settings to the selected shapes when \"Apply Format\" is used.";
            Func<string, string> setterToolTip = description => "Determines " + description + " the shape. Use \"Apply Format\" to apply the setting.";
            Func<string, string> checkboxToolTip = description => "Check if you want the " + description + " setting to be synced when using \"Apply Format\".";

            SetTooltip(ToggleSyncLineButton, toggleSyncButtonToolTip("line style"));
            SetTooltip(ToggleSyncFillButton, toggleSyncButtonToolTip("fill style"));
            SetTooltip(ToggleSyncTextButton, toggleSyncButtonToolTip("text format"));
            SetTooltip(ToggleSyncSizeButton, toggleSyncButtonToolTip("size"));

            SetTooltip(FormatHasLine, setterToolTip("whether line is enabled for"));
            SetTooltip(FormatIncludeHasLine, checkboxToolTip("has line"));
            SetTooltip(FormatLineColor, setterToolTip("the line color of"));
            SetTooltip(FormatIncludeLineColor, checkboxToolTip("line color"));
            SetTooltip(FormatLineWeight, setterToolTip("the line thickness of"));
            SetTooltip(FormatIncludeLineWeight, checkboxToolTip("line thickness"));
            SetTooltip(FormatLineDashStyle, setterToolTip("the dash style (e.g. solid, dashed) of"));
            SetTooltip(FormatIncludeLineDashStyle, checkboxToolTip("dash style"));

            SetTooltip(FormatHasFill, setterToolTip("whether fill is enabled for"));
            SetTooltip(FormatIncludeHasFill, checkboxToolTip("has fill"));
            SetTooltip(FormatFillColor, setterToolTip("the fill color of"));
            SetTooltip(FormatIncludeFillColor, checkboxToolTip("fill color"));

            SetTooltip(FormatText, setterToolTip("the text content of"));
            SetTooltip(FormatIncludeText, checkboxToolTip("text content"));
            SetTooltip(FormatTextColor, setterToolTip("the color of the text in"));
            SetTooltip(FormatIncludeTextColor, checkboxToolTip("text color"));
            SetTooltip(FormatTextFontSize, setterToolTip("the font size of the text in"));
            SetTooltip(FormatIncludeTextFontSize, checkboxToolTip("font size"));
            SetTooltip(FormatTextFont, setterToolTip("the font of the text in"));
            SetTooltip(FormatIncludeTextFont, checkboxToolTip("font"));
            SetTooltip(FormatTextWrap, setterToolTip("whether text wrap is enabled for"));
            SetTooltip(FormatIncludeTextWrap, checkboxToolTip("text wrap"));
            SetTooltip(FormatTextAutoSize, setterToolTip("the text auto size setting (e.g. shrink text to fit) for"));
            SetTooltip(FormatIncludeTextAutoSize, checkboxToolTip("text auto size"));

            SetTooltip(FormatWidth, setterToolTip("the width of"));
            SetTooltip(FormatIncludeWidth, checkboxToolTip("width"));
            SetTooltip(FormatHeight, setterToolTip("the height of"));
            SetTooltip(FormatIncludeHeight, checkboxToolTip("height"));


            // -------------------
            // || Tab: Position ||
            // -------------------

            configureButton(ApplyDisplacementButton, ()=>_drawingLab.ApplyDisplacement(applyAllSettings: false), "Apply recorded displacement to selected shapes.");
            configureButton(ApplyPositionButton, () => _drawingLab.ApplyPosition(applyAllSettings: false), "Apply recorded position or rotation to selected shapes.");
            configureButton(RecordDisplacementButton, _drawingLab.RecordDisplacement, "Record the Displacement between two selected shapes.");
            configureButton(RecordPositionButton, _drawingLab.RecordPosition, "Record position and rotation of a selected shape.");

            configureButton(AlignHorizontalButton, _drawingLab.AlignHorizontal, "Align Shapes Horizontally to last shape in selection.");
            configureButton(AlignVerticalButton, _drawingLab.AlignVertical, "Align Shapes Vertically to last shape in selection.");
            configureButton(AlignBothButton, _drawingLab.AlignBoth, "Align Shapes both Horizontally and Vertically to last shape in selection.");
            
            configureButton(PivotAroundButton, _drawingLab.PivotAroundTool, "Rotate / Multiclone a shape around another shape. Two shapes must be selected, the shape to be rotated and the pivot, in order.");

            configureButton(AlignHorizontalToSlideButton, _drawingLab.AlignHorizontalToSlide, "Align Shapes Horizontally to a position relative to the slide.");
            configureButton(AlignVerticalToSlideButton, _drawingLab.AlignVerticalToSlide, "Align Shapes Vertically to a position relative to the slide.");
            configureButton(AlignBothToSlideButton, _drawingLab.AlignBothToSlide, "Align Shapes both Horizontally and Vertically to a position relative to the slide.");

            Func<string, string> savedSetterToolTip = description => "The " + description + " of the shape. Use \"Apply Position\" to apply the setting.";
            Func<string, string> savedCheckboxToolTip = description => "Check if you want to apply the " + description + " when using \"Apply Position\".";
            Func<string, string> shiftSetterToolTip = description => "A difference in the " + description + ". Use \"Apply Displacement\" to shift the selected shapes by this amount.";
            Func<string, string> shiftCheckboxToolTip = description => "Check if you want to apply the " + description + " displacement when using \"Apply Displacement\".";

            SetTooltip(SavedValueX, savedSetterToolTip("x-coordinate"));
            SetTooltip(SavedIncludeX, savedCheckboxToolTip("x-coordinate position"));
            SetTooltip(SavedValueY, savedSetterToolTip("y-coordinate"));
            SetTooltip(SavedIncludeY, savedCheckboxToolTip("y-coordinate position"));
            SetTooltip(SavedValueRotation, savedSetterToolTip("orientation"));
            SetTooltip(SavedIncludeRotation, savedCheckboxToolTip("orientation"));

            SetTooltip(ShiftValueX, shiftSetterToolTip("x-coordinate"));
            SetTooltip(ShiftIncludeX, shiftCheckboxToolTip("x-coordinate"));
            SetTooltip(ShiftValueY, shiftSetterToolTip("y-coordinate"));
            SetTooltip(ShiftIncludeY, shiftCheckboxToolTip("y-coordinate"));
            SetTooltip(ShiftValueRotation, shiftSetterToolTip("orientation"));
            SetTooltip(ShiftIncludeRotation, shiftCheckboxToolTip("angular"));

            Func<string, string> anchorToolTip = description => "Set the anchor when Recording/Applying shape positions / displacements to the " + description + " of the shape.\nThe anchor is used as the reference point to determine the shape's coordinates.";

            SetTooltip(AnchorTopLeft, anchorToolTip("top left"));
            SetTooltip(AnchorTopCen, anchorToolTip("top center"));
            SetTooltip(AnchorTopRight, anchorToolTip("top right"));
            SetTooltip(AnchorMidLeft, anchorToolTip("middle left"));
            SetTooltip(AnchorMidCen, anchorToolTip("middle center"));
            SetTooltip(AnchorMidRight, anchorToolTip("middle right"));
            SetTooltip(AnchorBotLeft, anchorToolTip("bottom left"));
            SetTooltip(AnchorBotCen, anchorToolTip("bottom center"));
            SetTooltip(AnchorBotRight, anchorToolTip("bottom right"));

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
            if (s.StartsWith("VK_"))
            {
                s = s.Substring(3);
            }

            return s;
        }

        #endregion

        #region Hotkey Initialisation
        private bool IsPanelOpen()
        {
            var drawingsPane = this.GetTaskPane(typeof(DrawingsPane));
            return drawingsPane != null && drawingsPane.Visible;
        }

        private bool IsReadingHotkeys()
        {
            // Is reading hotkeys when panel is open and user is not selecting text.
            return IsPanelOpen() &&
                   _dataSource.HotkeysEnabled &&
                   this.GetCurrentSelection().Type != PpSelectionType.ppSelectionText;
        }

        private Action RunOnlyWhenOpen(Action action)
        {
            return () =>
            {
                if (IsReadingHotkeys())
                {
                    action();
                }
            };
        }

        private Dictionary<int, Native.VirtualKey> SetupButtonHotkeys()
        {
            var usedKeys = new HashSet<Native.VirtualKey>();
            var bindings = new Dictionary<int, Native.VirtualKey>();
            Action<Native.VirtualKey, ImageButton> assign = (key, button) =>
            {
                bindings.Add(button.ImageButtonUniqueId, key);
                if (usedKeys.Contains(key))
                {
                    throw new ArgumentException("Key already has a binding: " + key.ToString());
                }
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
            if (_data.IsHotkeysInitialised)
            {
                return;
            }

            _data.IsHotkeysInitialised = true;

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
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => _drawingLab.SelectControlGroup(key)));
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => _drawingLab.SetControlGroup(key)), ctrl: true);

                // Block shift+number and ctrl+shift+number
                PPKeyboard.AddConditionToBlockTextInput(IsReadingHotkeys, key, shift: true);
                PPKeyboard.AddConditionToBlockTextInput(IsReadingHotkeys, key, ctrl: true, shift: true);

                // Assign shift+number and ctrl+shift+number to control group commands.
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => _drawingLab.SelectControlGroup(key, appendToSelection: true)), shift: true);
                PPKeyboard.AddKeyupAction(key, RunOnlyWhenOpen(() => _drawingLab.SetControlGroup(key, appendToGroup: true)), ctrl: true, shift: true);
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
                Color = Graphics.ConvertRgbToColor(_dataSource.FormatFillColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            _dataSource.FormatFillColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }

        private void LineColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog
            {
                Color = Graphics.ConvertRgbToColor(_dataSource.FormatLineColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            _dataSource.FormatLineColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }

        private void TextColor_Click(object sender, EventArgs e)
        {
            var colorDialog = new ColorDialog
            {
                Color = Graphics.ConvertRgbToColor(_dataSource.FormatTextColor),
                FullOpen = true
            };
            if (colorDialog.ShowDialog() == DialogResult.Cancel)
            {
                return;
            }

            _dataSource.FormatTextColor = Graphics.ConvertColorToRgb(colorDialog.Color);
        }
    }
}
