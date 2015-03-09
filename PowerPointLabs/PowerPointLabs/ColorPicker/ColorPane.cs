using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace PowerPointLabs
{

    public partial class ColorPane : UserControl
    {
        // To set eyedropper mode
        private enum MODE
        {
            FILL,
            LINE,
            FONT,
            NONE
        };

        // Keeps track of mouse on mouse down on a matching panel.
        // Needed to determine drag-drop v/s click
        private System.Drawing.Point _mouseDownLocation;
        
        // Listener for Mouse Up Events (for EyeDropping)
        LMouseUpListener _native;

        // Shapes to Update
        PowerPoint.ShapeRange _selectedShapes;

        PowerPoint.TextRange _selectedText;
        
        // Data-bindings datasource
        ColorDataSource dataSource = new ColorDataSource();

        private Color _pickedColor;

        // Stores last selected mode
        private MODE prevMode = MODE.NONE;
        
        // Stores the current selected mode
        private MODE currMode = MODE.NONE;

        private String _defaultThemeColorDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "PowerPointLabs.defaultThemeColor.thm");

        public ColorPane()
        {
            InitializeComponent();

            BindDataToPanels();

            InitToolTipControl();

            LoadThemePanel();

            // Default color to CornFlowerBlue
            SetDefaultColor(Color.CornflowerBlue);
        }

        public void SaveDefaultColorPaneThemeColors()
        {
            dataSource.SaveThemeColorsInFile(_defaultThemeColorDirectory);
        }

        private void SetDefaultColor(Color color)
        {
            dataSource.selectedColor = color;
            UpdateUIForNewColor();
        }

        private void SetMode(MODE mode)
        {
            dataSource.isFillColorSelected = false;
            dataSource.isFontColorSelected = false;
            dataSource.isLineColorSelected = false;
            
            switch (mode)
            {
                case MODE.LINE:
                    dataSource.isLineColorSelected = true;
                    break;
                case MODE.FONT:
                    dataSource.isFontColorSelected = true;
                    break;
                case MODE.FILL:
                    dataSource.isFillColorSelected = true;
                    break;
                default:
                    currMode = MODE.NONE;
                    break;
            }
            UpdateCurrMode();
        }

        private void ResetEyeDropperSelectionInDataSource()
        {
            dataSource.isFillColorSelected = false;
            dataSource.isFontColorSelected = false;
            dataSource.isLineColorSelected = false;
        }

        #region ToolTip
        private void InitToolTipControl()
        {
            toolTip1.SetToolTip(panel1, TextCollection.ColorsLabText.MainColorBoxTooltips);
            toolTip1.SetToolTip(this.FontButton, TextCollection.ColorsLabText.FontColorButtonTooltips);
            toolTip1.SetToolTip(this.LineButton, TextCollection.ColorsLabText.LineColorButtonTooltips);
            toolTip1.SetToolTip(this.FillButton, TextCollection.ColorsLabText.FillColorButtonTooltips);
            toolTip1.SetToolTip(panel2, TextCollection.ColorsLabText.BrightnessSliderTooltips);
            toolTip1.SetToolTip(panel3, TextCollection.ColorsLabText.SaturationSliderTooltips);
            toolTip1.SetToolTip(this.SaveThemeButton, TextCollection.ColorsLabText.SaveFavoriteColorsButtonTooltips);
            toolTip1.SetToolTip(this.LoadButton, TextCollection.ColorsLabText.LoadFavoriteColorsButtonTooltips);
            toolTip1.SetToolTip(this.ResetThemeButton, TextCollection.ColorsLabText.ResetFavoriteColorsButtonTooltips);
            toolTip1.SetToolTip(this.EmptyPanelButton, TextCollection.ColorsLabText.EmptyFavoriteColorsButtonTooltips);
            const string colorRectangleToolTip = TextCollection.ColorsLabText.ColorRectangleTooltips;
            const string themeColorRectangleToolTip = TextCollection.ColorsLabText.ThemeColorRectangleTooltips;
            toolTip1.SetToolTip(this.ThemePanel1, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel2, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel3, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel4, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel5, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel6, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel7, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel8, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel9, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.ThemePanel10, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.MonoPanel1, colorRectangleToolTip);
            toolTip1.SetToolTip(this.MonoPanel2, colorRectangleToolTip);
            toolTip1.SetToolTip(this.MonoPanel3, colorRectangleToolTip);
            toolTip1.SetToolTip(this.MonoPanel4, colorRectangleToolTip);
            toolTip1.SetToolTip(this.MonoPanel5, colorRectangleToolTip);
            toolTip1.SetToolTip(this.MonoPanel6, colorRectangleToolTip);
            toolTip1.SetToolTip(this.MonoPanel7, colorRectangleToolTip);
            toolTip1.SetToolTip(this.AnalogousLighter, colorRectangleToolTip);
            toolTip1.SetToolTip(this.AnalogousSelected, colorRectangleToolTip);
            toolTip1.SetToolTip(this.AnalogousDarker, colorRectangleToolTip);
            toolTip1.SetToolTip(this.ComplementaryLighter, colorRectangleToolTip);
            toolTip1.SetToolTip(this.ComplementarySelected, colorRectangleToolTip);
            toolTip1.SetToolTip(this.ComplementaryDarker, colorRectangleToolTip);
            toolTip1.SetToolTip(this.TriadicLower, colorRectangleToolTip);
            toolTip1.SetToolTip(this.TriadicSelected, colorRectangleToolTip);
            toolTip1.SetToolTip(this.TriadicHigher, colorRectangleToolTip);
            toolTip1.SetToolTip(this.Tetradic1, colorRectangleToolTip);
            toolTip1.SetToolTip(this.Tetradic2, colorRectangleToolTip);
            toolTip1.SetToolTip(this.Tetradic3, colorRectangleToolTip);
            toolTip1.SetToolTip(this.TetradicSelected, colorRectangleToolTip);
        }
        #endregion

        #region DataBindings
        private void BindDataToPanels()
        {
            this.panel1.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel1.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorOne",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel2.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorTwo",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel3.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorThree",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel4.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorFour",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel5.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorFive",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel6.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorSix",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel7.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorSeven",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel8.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorEight",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel9.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorNine",
                new Converters.HSLColorToRGBColor()));

            this.ThemePanel10.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorTen",
                new Converters.HSLColorToRGBColor()));

            this.AnalogousSelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.HSLColorToRGBColor()));

            this.AnalogousLighter.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToAnalogousLower()));

            this.AnalogousDarker.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToAnalogousHigher()));

            this.ComplementarySelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToComplementaryColor()));

            this.ComplementaryLighter.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToSplitComplementaryLower()));

            this.ComplementaryDarker.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToSplitComplementaryHigher()));

            this.TriadicSelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.HSLColorToRGBColor()));

            this.TriadicLower.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToTriadicLower()));

            this.TriadicHigher.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToTriadicHigher()));

            this.TetradicSelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.HSLColorToRGBColor()));

            this.Tetradic1.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.selectedColorToTetradicOne()));

            this.Tetradic2.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToTetradicTwo()));

            this.Tetradic3.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToTetradicThree()));

            this.MonoPanel1.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToMonochromaticOne()));

            this.MonoPanel2.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToMonochromaticTwo()));

            this.MonoPanel3.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToMonochromaticThree()));

            this.MonoPanel4.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToMonochromaticFour()));

            this.MonoPanel5.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToMonochromaticFive()));

            this.MonoPanel6.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToMonochromaticSix()));

            this.MonoPanel7.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToMonochromaticSeven()));

            brightnessBar.DataBindings.Add(new CustomBinding(
                            "Value",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToBrightnessValue()));

            saturationBar.DataBindings.Add(new CustomBinding(
                            "Value",
                            dataSource,
                            "selectedColor",
                            new Converters.selectedColorToSaturationValue()));

            FillButton.DataBindings.Add(new CustomBinding(
                        "BackColor",
                        dataSource,
                        "isFillColorSelected",
                        new Converters.IsActiveBoolToButtonBackColorConverter()));
            
            LineButton.DataBindings.Add(new CustomBinding(
                        "BackColor",
                        dataSource,
                        "isLineColorSelected",
                        new Converters.IsActiveBoolToButtonBackColorConverter()));
            
            FontButton.DataBindings.Add(new CustomBinding(
                        "BackColor",
                        dataSource,
                        "isFontColorSelected",
                        new Converters.IsActiveBoolToButtonBackColorConverter()));
        }

        #endregion

        private int _timerCounter = 0;
        private const int TIMER_COUNTER_THRESHOLD = 2;
        private Cursor eyeDropperCursor = new Cursor(new MemoryStream(Properties.Resources.EyeDropper));

        private void BeginEyedropping()
        {
            if (!VerifyIsShapeSelected()) return;

            _timerCounter = 0;
            timer1.Start();
            _native = new LMouseUpListener();
            _native.LButtonUpClicked +=
                new EventHandler<SysMouseEventInfo>(_native_LButtonClicked);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //this is to ensure that EyeDropper tool feature doesn't
            //affect Color Dialog tool feature
            if (_timerCounter < TIMER_COUNTER_THRESHOLD)
            {
                _timerCounter++;
                return;
            }
            _timerCounter++;

            Cursor.Current = eyeDropperCursor;
            System.Drawing.Point mousePos = Control.MousePosition;
            IntPtr deviceContext = Native.GetDC(IntPtr.Zero);
            if (currMode == MODE.NONE)
            {
                dataSource.selectedColor = ColorTranslator.FromWin32(Native.GetPixel(deviceContext, mousePos.X, mousePos.Y));
                ColorSelectedShapesWithColor(panel1.BackColor);
            }
            else
            {
                _pickedColor = ColorTranslator.FromWin32(Native.GetPixel(deviceContext, mousePos.X, mousePos.Y));
                ColorSelectedShapesWithColor(_pickedColor);
            }
        }

        #region Selection And Coloring Shapes
        private void SelectShapes()
        {
            try
            {
                var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
                if (selection == null) return;

                if (selection.Type == PpSelectionType.ppSelectionShapes &&
                    selection.HasChildShapeRange)
                {
                    _selectedShapes = selection.ChildShapeRange;
                }
                else if (selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    _selectedShapes = selection.ShapeRange;
                } else if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    _selectedText = selection.TextRange;
                }
                else
                {
                    _selectedShapes = null;
                    _selectedText = null;
                }
            }
            catch (Exception)
            {
                _selectedShapes = null;
                _selectedText = null;
            }
        }
        private void ColorSelectedShapesWithColor(HSLColor selectedColor)
        {
            SelectShapes();
            if (_selectedShapes != null
                && PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape s in _selectedShapes)
                {
                    try
                    {
                        var r = ((Color)selectedColor).R;
                        var g = ((Color)selectedColor).G;
                        var b = ((Color)selectedColor).B;

                        var rgb = (b << 16) | (g << 8) | (r);
                        ColorShapeWithColor(s, rgb, currMode);
                    }
                    catch (Exception)
                    {
                        RecreateCorruptedShape(s);
                    }  
                }
            }
            if (_selectedText != null
                && PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionText)
            {
                try
                {
                    var r = ((Color)selectedColor).R;
                    var g = ((Color)selectedColor).G;
                    var b = ((Color)selectedColor).B;

                    var rgb = (b << 16) | (g << 8) | (r);
                    ColorShapeWithColor(_selectedText, rgb, currMode);
                }
                catch (Exception e)
                {
                }  
            }
        }

        private static void RecreateCorruptedShape(PowerPoint.Shape s)
        {
            s.Copy();
            PowerPoint.Shape newShape = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.Paste()[1];

            newShape.Select();

            newShape.Name = s.Name;
            newShape.Left = s.Left;
            newShape.Top = s.Top;
            while (newShape.ZOrderPosition > s.ZOrderPosition)
            {
                newShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBackward);
            }
            s.Delete();
        }

        private void ColorShapeWithColor(PowerPoint.Shape s, int rgb, MODE mode)
        {
            switch (mode)
            {
                case MODE.FILL:
                    s.Fill.ForeColor.RGB = rgb;
                    break;
                case MODE.LINE:
                    s.Line.ForeColor.RGB = rgb;
                    s.Line.Visible = MsoTriState.msoTrue;
                    break;
                case MODE.FONT:
                    ColorShapeFontWithColor(s, rgb);
                    break;
            }
        }

        private static void ColorShapeFontWithColor(PowerPoint.Shape s, int rgb)
        {
            if (s.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.HasTextFrame
                    == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    if (PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionText)
                    {
                        TextRange selectedText
                            = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.TrimText();
                        if (selectedText.Text != "" && selectedText != null)
                        {
                            selectedText.Font.Color.RGB = rgb;
                        }
                        else
                        {
                            s.TextFrame.TextRange.TrimText().Font.Color.RGB = rgb;
                        }
                    }
                    else if (PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionShapes)
                    {
                        s.TextFrame.TextRange.TrimText().Font.Color.RGB = rgb;
                    }
                }
            }
        }

        private void ColorShapeWithColor(PowerPoint.TextRange text, int rgb, MODE mode)
        {
            var frame = text.Parent as PowerPoint.TextFrame;
            var selectedShape = frame.Parent as PowerPoint.Shape;
            if (mode != MODE.NONE)
            {
                ColorShapeWithColor(selectedShape, rgb, mode);
            }
        }

        void _native_LButtonClicked(object sender, SysMouseEventInfo e)
        {
            _native.Close();
            timer1.Stop();
            //this is to ensure that EyeDropper tool feature doesn't
            //affect Color Dialog tool feature
            if (_timerCounter >= TIMER_COUNTER_THRESHOLD)
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                UpdateUIForNewColor();
                if (currMode != MODE.NONE)
                {
                    ColorSelectedShapesWithColor(_pickedColor);
                }
            }
            _timerCounter = 0;
            ResetEyeDropperSelectionInDataSource();
        }
        #endregion

        private void UpdateCurrMode()
        {
            if (currMode != MODE.NONE)
            {
                prevMode = currMode;
            }

            if (dataSource.isFillColorSelected)
            {
                currMode = MODE.FILL;
            }
            else if (dataSource.isFontColorSelected)
            {
                currMode = MODE.FONT;
            }
            else if (dataSource.isLineColorSelected)
            {
                currMode = MODE.LINE;
            }
        }

        private void UpdateUIForNewColor()
        {
            UpdateBrightnessBar(dataSource.selectedColor);
            UpdateSaturationBar(dataSource.selectedColor);
        }

        #region Brightness and Saturation
        private void UpdateBrightnessBar(HSLColor color)
        {
            DrawBrightnessGradient(color);
        }

        private void brightnessBar_ValueChanged(object sender, EventArgs e)
        {
            if (!timer1.Enabled)
            {
                float newBrightness = brightnessBar.Value;
                var newColor = new HSLColor();
                try
                {
                    newColor.Hue = dataSource.selectedColor.Hue;
                    newColor.Saturation = dataSource.selectedColor.Saturation;
                    newColor.Luminosity = newBrightness;

                    brightnessBar.ValueChanged -= brightnessBar_ValueChanged;
                    saturationBar.ValueChanged -= saturationBar_ValueChanged;

                    dataSource.selectedColor = newColor;
                    UpdateSaturationBar(newColor);
                    UpdateBrightnessBar(newColor);

                    brightnessBar.ValueChanged += brightnessBar_ValueChanged;
                    saturationBar.ValueChanged += saturationBar_ValueChanged;
                }
                catch (Exception)
                {
                }
            }
        }

        private void saturationBar_ValueChanged(object sender, EventArgs e)
        {
            float newSaturation = saturationBar.Value;
            var newColor = new HSLColor();
            try
            {
                newColor.Hue = dataSource.selectedColor.Hue;
                newColor.Saturation = newSaturation;
                newColor.Luminosity = dataSource.selectedColor.Luminosity;

                brightnessBar.ValueChanged -= brightnessBar_ValueChanged;
                saturationBar.ValueChanged -= saturationBar_ValueChanged;

                dataSource.selectedColor = newColor;
                UpdateBrightnessBar(newColor);
                UpdateSaturationBar(newColor);
                
                brightnessBar.ValueChanged += brightnessBar_ValueChanged;
                saturationBar.ValueChanged += saturationBar_ValueChanged;
            }
            catch (Exception)
            {
            }
        }
        private void UpdateSaturationBar(HSLColor color)
        {
            DrawSaturationGradient(color);
        }

        private void DrawBrightnessGradient(HSLColor color)
        {
            var dis = brightnessPanel.DisplayRectangle;
            var screenRec = brightnessPanel.RectangleToScreen(dis);
            var rec = brightnessPanel.Parent.RectangleToClient(screenRec);
            brightnessPanel.Visible = false;
            LinearGradientBrush brush = new LinearGradientBrush(
                rec,
                Color.Transparent,
                Color.Transparent,
                LinearGradientMode.Horizontal);
            ColorBlend blend = new ColorBlend();
            Color[] blendColors = {
                new HSLColor(
                color.Hue,
                color.Saturation,
                0),
                new HSLColor(
                color.Hue,
                color.Saturation,
                120),
                new HSLColor(
                color.Hue,
                color.Saturation,
                240)};
            float[] positions = { 0.0f, 0.5f, 1.0f };
            blend.Colors = blendColors;
            blend.Positions = positions;

            brush.InterpolationColors = blend;

            using (Graphics g = this.CreateGraphics())
            {
                g.FillRectangle(brush, rec);
            }
        }

        private void DrawSaturationGradient(HSLColor color)
        {
            var dis = saturationPanel.DisplayRectangle;
            var screenRec = saturationPanel.RectangleToScreen(dis);
            var rec = saturationPanel.Parent.RectangleToClient(screenRec);
            saturationPanel.Visible = false;
            LinearGradientBrush brush = new LinearGradientBrush(
                rec,
                Color.Transparent,
                Color.Transparent,
                LinearGradientMode.Horizontal);
            ColorBlend blend = new ColorBlend();
            Color[] blendColors = {
                new HSLColor(
                color.Hue,
                0,
                color.Luminosity),
                new HSLColor(
                color.Hue,
                120,
                color.Luminosity),
                new HSLColor(
                color.Hue,
                240,
                color.Luminosity)};
            float[] positions = { 0.0f, 0.5f, 1.0f };
            blend.Colors = blendColors;
            blend.Positions = positions;

            brush.InterpolationColors = blend;

            using (Graphics g = this.CreateGraphics())
            {
                g.FillRectangle(brush, rec);
            }
        }

        protected override void OnPaint(PaintEventArgs paintEvnt)
        {
            DrawBrightnessGradient(dataSource.selectedColor);
            DrawSaturationGradient(dataSource.selectedColor);
        }

        #endregion

        #region Drag-Drop
        private Cursor openHandCursor = new Cursor(new MemoryStream(Properties.Resources.OpenHand));
        private Cursor closedHandCursor = new Cursor(new MemoryStream(Properties.Resources.ClosedHand));

        private void EyeDropButton_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button)
            {
                Button button = sender as Button;
                button.Cursor = openHandCursor;
            } else if (sender is Panel)
            {
                Panel panel = sender as Panel;
                panel.Cursor = openHandCursor;
            }
        }

        private void MatchingPanel_MouseEnter(object sender, EventArgs e)
        {
            Panel panel = sender as Panel;
            panel.Cursor = openHandCursor;
        }

        private void MatchingPanel_MouseUp(object sender, MouseEventArgs e)
        {
            Panel panel = sender as Panel;
            panel.Cursor = openHandCursor;
        }

        private void MatchingPanel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            _mouseDownLocation = e.Location;
            Panel panel = sender as Panel;
            panel.Cursor = closedHandCursor;
        }

        private void StartDragDrop(object sender)
        {
            DataObject colorObject = new DataObject();
            colorObject.SetData(((Panel)sender).BackColor);
            DoDragDrop(colorObject, DragDropEffects.All);
        }

        private void panel_DragDrop(object sender, DragEventArgs e)
        {
            Panel panel = (Panel)sender;
            panel.BackColor = (Color)e.Data.GetData(panel.BackColor.GetType());
            if (panel.Equals(panel1))
            {
                dataSource.selectedColor = panel.BackColor;
            }
            if (panel.Equals(ThemePanel1))
            {
                dataSource.themeColorOne = panel.BackColor;
            }
            if (panel.Equals(ThemePanel2))
            {
                dataSource.themeColorTwo = panel.BackColor;
            }
            if (panel.Equals(ThemePanel3))
            {
                dataSource.themeColorThree = panel.BackColor;
            }
            if (panel.Equals(ThemePanel4))
            {
                dataSource.themeColorFour = panel.BackColor;
            }
            if (panel.Equals(ThemePanel5))
            {
                dataSource.themeColorFive = panel.BackColor;
            }
            if (panel.Equals(ThemePanel6))
            {
                dataSource.themeColorSix = panel.BackColor;
            }
            if (panel.Equals(ThemePanel7))
            {
                dataSource.themeColorSeven = panel.BackColor;
            }
            if (panel.Equals(ThemePanel8))
            {
                dataSource.themeColorEight = panel.BackColor;
            }
            if (panel.Equals(ThemePanel9))
            {
                dataSource.themeColorNine = panel.BackColor;
            }
            if (panel.Equals(ThemePanel10))
            {
                dataSource.themeColorTen = panel.BackColor;
            }
            UpdateUIForNewColor();
        }

        private void panel_DragEnter(object sender, DragEventArgs e)
        {
            String colorTypeName = Color.Red.GetType().ToString();
            if (e.Data.GetDataPresent(colorTypeName))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void MatchingPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;

            int dx = e.X - _mouseDownLocation.X;
            int dy = e.Y - _mouseDownLocation.Y;

            if (Math.Abs(dx) > ((Panel)sender).Width / 2 ||
                Math.Abs(dy) > ((Panel)sender).Height / 2)
            {
                StartDragDrop(sender);
            }
        }
        #endregion

        #region Theme Functions
        private void LoadThemePanel()
        {
            Boolean isSuccessful = dataSource.LoadThemeColorsFromFile(_defaultThemeColorDirectory);
            if (!isSuccessful)
            {
                EmptyThemePanel();
            }
        }

        private void SaveThemeButton_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK && 
                dataSource.SaveThemeColorsInFile(saveFileDialog1.FileName))
            {
                SaveDefaultColorPaneThemeColors();
            }
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK &&
                dataSource.LoadThemeColorsFromFile(openFileDialog1.FileName))
            {
                SaveDefaultColorPaneThemeColors();
            }
        }

        private void MatchingPanel_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)sender).BackColor;
            dataSource.selectedColor = clickedColor;
            Globals.ThisAddIn.Application.StartNewUndoEntry();
        }

        private void ResetThemeButton_Click(object sender, EventArgs e)
        {
            ResetThemePanel();
        }

        private void ResetThemePanel()
        {
            try
            {
                LoadThemePanel();
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Theme Panel Reset Failed", e.Message, e);
            }
        }

        private void EmptyThemePanel()
        {
            try
            {
                if (PowerPointPresentation.Current.SlideCount > 0)
                {
                    dataSource.themeColorOne = Color.White;
                    dataSource.themeColorTwo = Color.White;
                    dataSource.themeColorThree = Color.White;
                    dataSource.themeColorFour = Color.White;
                    dataSource.themeColorFive = Color.White;
                    dataSource.themeColorSix = Color.White;
                    dataSource.themeColorSeven = Color.White;
                    dataSource.themeColorEight = Color.White;
                    dataSource.themeColorNine = Color.White;
                    dataSource.themeColorTen = Color.White;
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Theme Panel Reset Failed", e.Message, e);
            }
        }

        private void EmptyPanelButton_Click(object sender, EventArgs e)
        {
            EmptyThemePanel();
        }

        #endregion

        #region Context Menu Clicks
        private void showMoreInformationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)(contextMenuStrip1.SourceControl)).BackColor;
            ColorInformationDialog dialog = new ColorInformationDialog(clickedColor);
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.Show();
        }

        private void selectAsMainColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)contextMenuStrip1.SourceControl).BackColor;

            dataSource.selectedColor = clickedColor;
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            UpdateUIForNewColor();
        }
        #endregion

        private void EyeDropButton_MouseClick(object sender, MouseEventArgs e)
        {
            string buttonName = "";
            if (sender is Button)
            {
                buttonName = ((Button)sender).Name;
            }
            else if (sender is Panel)
            {
                buttonName = ((Panel)sender).Name;
            }
            SetModeForSenderName(buttonName);

            if (!VerifyIsShapeSelected()) return;

            colorDialog1.Color = GetSelectedShapeColor();

            DialogResult result = colorDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (currMode == MODE.NONE)
                {
                    SetDefaultColor(colorDialog1.Color);
                }
                else
                {
                    ColorSelectedShapesWithColor(colorDialog1.Color);
                }
                ResetEyeDropperSelectionInDataSource();
            }
            else if (result == DialogResult.Cancel)
            {
                ResetEyeDropperSelectionInDataSource();
            }
        }

        private Color GetSelectedShapeColor()
        {
            SelectShapes();
            if (_selectedShapes == null && _selectedText == null)
                return dataSource.selectedColor;

            if (PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionShapes)
            {
                return GetSelectedShapeColor(_selectedShapes);
            }
            if (PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionText)
            {
                var frame = _selectedText.Parent as PowerPoint.TextFrame;
                var selectedShape = frame.Parent as PowerPoint.Shape;
                return GetSelectedShapeColor(selectedShape);
            }

            return dataSource.selectedColor;
        }

        private Color GetSelectedShapeColor(PowerPoint.ShapeRange selectedShapes)
        {
            Color colorToReturn = Color.Empty;
            foreach (var selectedShape in selectedShapes)
            {
                Color color = GetSelectedShapeColor(selectedShape as PowerPoint.Shape);
                if (colorToReturn.Equals(Color.Empty))
                {
                    colorToReturn = color;
                }
                else if(!colorToReturn.Equals(color))
                {
                    return Color.Black;
                }
            }
            return colorToReturn;
        }

        private Color GetSelectedShapeColor(PowerPoint.Shape selectedShape)
        {
            switch (currMode)
            {
                case MODE.FILL:
                    return Color.FromArgb(ColorHelper.ReverseRGBToArgb(selectedShape.Fill.ForeColor.RGB));
                    break;
                case MODE.LINE:
                    return Color.FromArgb(ColorHelper.ReverseRGBToArgb(selectedShape.Line.ForeColor.RGB));
                    break;
                case MODE.FONT:
                    if (selectedShape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue
                        && Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.HasTextFrame
                        == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        if (PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionText)
                        {
                            var selectedText
                                = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.TrimText();
                            if (selectedText != null && selectedText.Text != "")
                            {
                                return Color.FromArgb(ColorHelper.ReverseRGBToArgb(selectedText.Font.Color.RGB));
                            }
                            else
                            {
                                return
                                Color.FromArgb(
                                    ColorHelper.ReverseRGBToArgb(selectedShape.TextFrame.TextRange.Font.Color.RGB));
                            }
                        }
                        else if (PowerPointCurrentPresentationInfo.CurrentSelection.Type == PpSelectionType.ppSelectionShapes)
                        {
                            return
                                Color.FromArgb(
                                    ColorHelper.ReverseRGBToArgb(selectedShape.TextFrame.TextRange.Font.Color.RGB));
                        }
                    }
                    break;
            }
            return dataSource.selectedColor;
        }

        private void SetModeForSenderName(string buttonName)
        {
            switch (buttonName)
            {
                case "FillButton":
                    SetMode(MODE.FILL);
                    break;
                case "FontButton":
                    SetMode(MODE.FONT);
                    break;
                case "LineButton":
                    SetMode(MODE.LINE);
                    break;
                default:
                    SetMode(MODE.NONE);
                    break;
            }
        }

        private void EyeDropButton_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;

            string buttonName = "";
            if (sender is Button)
            {
                buttonName = ((Button)sender).Name;
            } 
            else if (sender is Panel)
            {
                buttonName = ((Panel)sender).Name;
            }
            SetModeForSenderName(buttonName);
            BeginEyedropping();
        }

        private Boolean VerifyIsShapeSelected()
        {
            SelectShapes();
            if (_selectedShapes == null && _selectedText == null && currMode != MODE.NONE)
            {
                MessageBox.Show(TextCollection.ColorsLabText.InfoHowToActivateFeature, "Colors Lab");
                return false;
            }
            return true;
        }

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
