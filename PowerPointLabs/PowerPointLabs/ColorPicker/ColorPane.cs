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
#pragma warning disable 0618
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

            EnableScrolling();
        }

        private void EnableScrolling()
        {
            // Account for scrollbar width
            this.Width += System.Windows.Forms.SystemInformation.VerticalScrollBarWidth;

            // Prevent horizontal scrollbar from appearing
            this.HorizontalScroll.Visible = false;
            this.HorizontalScroll.Maximum = 0;

            // Enable Autoscroll
            this.AutoScroll = true;
        }

        public void SaveDefaultColorPaneThemeColors()
        {
            dataSource.SaveThemeColorsInFile(_defaultThemeColorDirectory);
        }

        private void SetDefaultColor(Color color)
        {
            dataSource.SelectedColor = color;
            UpdateUIForNewColor();
        }

        private void SetMode(MODE mode)
        {
            dataSource.IsFillColorSelected = false;
            dataSource.IsFontColorSelected = false;
            dataSource.IsLineColorSelected = false;
            
            switch (mode)
            {
                case MODE.LINE:
                    dataSource.IsLineColorSelected = true;
                    break;
                case MODE.FONT:
                    dataSource.IsFontColorSelected = true;
                    break;
                case MODE.FILL:
                    dataSource.IsFillColorSelected = true;
                    break;
                default:
                    currMode = MODE.NONE;
                    break;
            }
            UpdateCurrMode();
        }

        private void ResetEyeDropperSelectionInDataSource()
        {
            dataSource.IsFillColorSelected = false;
            dataSource.IsFontColorSelected = false;
            dataSource.IsLineColorSelected = false;
        }

        #region ToolTip
        private void InitToolTipControl()
        {
            toolTip1.SetToolTip(panel1, TextCollection.ColorsLabText.MainColorBoxTooltips);
            toolTip1.SetToolTip(this.fontButton, TextCollection.ColorsLabText.FontColorButtonTooltips);
            toolTip1.SetToolTip(this.lineButton, TextCollection.ColorsLabText.LineColorButtonTooltips);
            toolTip1.SetToolTip(this.fillButton, TextCollection.ColorsLabText.FillColorButtonTooltips);
            toolTip1.SetToolTip(panel2, TextCollection.ColorsLabText.BrightnessSliderTooltips);
            toolTip1.SetToolTip(panel3, TextCollection.ColorsLabText.SaturationSliderTooltips);
            toolTip1.SetToolTip(this.saveThemeButton, TextCollection.ColorsLabText.SaveFavoriteColorsButtonTooltips);
            toolTip1.SetToolTip(this.loadButton, TextCollection.ColorsLabText.LoadFavoriteColorsButtonTooltips);
            toolTip1.SetToolTip(this.resetThemeButton, TextCollection.ColorsLabText.ResetFavoriteColorsButtonTooltips);
            toolTip1.SetToolTip(this.emptyPanelButton, TextCollection.ColorsLabText.EmptyFavoriteColorsButtonTooltips);
            const string colorRectangleToolTip = TextCollection.ColorsLabText.ColorRectangleTooltips;
            const string themeColorRectangleToolTip = TextCollection.ColorsLabText.ThemeColorRectangleTooltips;
            toolTip1.SetToolTip(this.themePanel1, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel2, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel3, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel4, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel5, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel6, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel7, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel8, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel9, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.themePanel10, themeColorRectangleToolTip);
            toolTip1.SetToolTip(this.monoPanel1, colorRectangleToolTip);
            toolTip1.SetToolTip(this.monoPanel2, colorRectangleToolTip);
            toolTip1.SetToolTip(this.monoPanel3, colorRectangleToolTip);
            toolTip1.SetToolTip(this.monoPanel4, colorRectangleToolTip);
            toolTip1.SetToolTip(this.monoPanel5, colorRectangleToolTip);
            toolTip1.SetToolTip(this.monoPanel6, colorRectangleToolTip);
            toolTip1.SetToolTip(this.monoPanel7, colorRectangleToolTip);
            toolTip1.SetToolTip(this.analogousLighter, colorRectangleToolTip);
            toolTip1.SetToolTip(this.analogousSelected, colorRectangleToolTip);
            toolTip1.SetToolTip(this.analogousDarker, colorRectangleToolTip);
            toolTip1.SetToolTip(this.complementaryLighter, colorRectangleToolTip);
            toolTip1.SetToolTip(this.complementarySelected, colorRectangleToolTip);
            toolTip1.SetToolTip(this.complementaryDarker, colorRectangleToolTip);
            toolTip1.SetToolTip(this.triadicLower, colorRectangleToolTip);
            toolTip1.SetToolTip(this.triadicSelected, colorRectangleToolTip);
            toolTip1.SetToolTip(this.triadicHigher, colorRectangleToolTip);
            toolTip1.SetToolTip(this.tetradic1, colorRectangleToolTip);
            toolTip1.SetToolTip(this.tetradic2, colorRectangleToolTip);
            toolTip1.SetToolTip(this.tetradic3, colorRectangleToolTip);
            toolTip1.SetToolTip(this.tetradicSelected, colorRectangleToolTip);
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

            this.themePanel1.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorOne",
                new Converters.HSLColorToRGBColor()));

            this.themePanel2.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorTwo",
                new Converters.HSLColorToRGBColor()));

            this.themePanel3.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorThree",
                new Converters.HSLColorToRGBColor()));

            this.themePanel4.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorFour",
                new Converters.HSLColorToRGBColor()));

            this.themePanel5.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorFive",
                new Converters.HSLColorToRGBColor()));

            this.themePanel6.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorSix",
                new Converters.HSLColorToRGBColor()));

            this.themePanel7.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorSeven",
                new Converters.HSLColorToRGBColor()));

            this.themePanel8.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorEight",
                new Converters.HSLColorToRGBColor()));

            this.themePanel9.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorNine",
                new Converters.HSLColorToRGBColor()));

            this.themePanel10.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "themeColorTen",
                new Converters.HSLColorToRGBColor()));

            this.analogousSelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.HSLColorToRGBColor()));

            this.analogousLighter.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToAnalogousLower()));

            this.analogousDarker.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToAnalogousHigher()));

            this.complementarySelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToComplementaryColor()));

            this.complementaryLighter.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToSplitComplementaryLower()));

            this.complementaryDarker.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToSplitComplementaryHigher()));

            this.triadicSelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.HSLColorToRGBColor()));

            this.triadicLower.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToTriadicLower()));

            this.triadicHigher.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToTriadicHigher()));

            this.tetradicSelected.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.HSLColorToRGBColor()));

            this.tetradic1.DataBindings.Add(new CustomBinding(
                "BackColor",
                dataSource,
                "selectedColor",
                new Converters.SelectedColorToTetradicOne()));

            this.tetradic2.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToTetradicTwo()));

            this.tetradic3.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToTetradicThree()));

            this.monoPanel1.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToMonochromaticOne()));

            this.monoPanel2.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToMonochromaticTwo()));

            this.monoPanel3.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToMonochromaticThree()));

            this.monoPanel4.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToMonochromaticFour()));

            this.monoPanel5.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToMonochromaticFive()));

            this.monoPanel6.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToMonochromaticSix()));

            this.monoPanel7.DataBindings.Add(new CustomBinding(
                            "BackColor",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToMonochromaticSeven()));

            brightnessBar.DataBindings.Add(new CustomBinding(
                            "Value",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToBrightnessValue()));

            saturationBar.DataBindings.Add(new CustomBinding(
                            "Value",
                            dataSource,
                            "selectedColor",
                            new Converters.SelectedColorToSaturationValue()));

            fillButton.DataBindings.Add(new CustomBinding(
                        "BackColor",
                        dataSource,
                        "isFillColorSelected",
                        new Converters.IsActiveBoolToButtonBackColorConverter()));
            
            lineButton.DataBindings.Add(new CustomBinding(
                        "BackColor",
                        dataSource,
                        "isLineColorSelected",
                        new Converters.IsActiveBoolToButtonBackColorConverter()));
            
            fontButton.DataBindings.Add(new CustomBinding(
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
            PPMouse.LeftButtonUp += LeftMouseButtonUpEventHandler;
        }

        private void Timer1_Tick(object sender, EventArgs e)
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
                dataSource.SelectedColor = ColorTranslator.FromWin32(Native.GetPixel(deviceContext, mousePos.X, mousePos.Y));
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
                }
                else if (selection.Type == PpSelectionType.ppSelectionText)
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
                catch (Exception)
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

        void LeftMouseButtonUpEventHandler()
        {
            PPMouse.LeftButtonUp -= LeftMouseButtonUpEventHandler;
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

            if (dataSource.IsFillColorSelected)
            {
                currMode = MODE.FILL;
            }
            else if (dataSource.IsFontColorSelected)
            {
                currMode = MODE.FONT;
            }
            else if (dataSource.IsLineColorSelected)
            {
                currMode = MODE.LINE;
            }
        }

        private void UpdateUIForNewColor()
        {
            UpdateBrightnessBar(dataSource.SelectedColor);
            UpdateSaturationBar(dataSource.SelectedColor);
        }

        #region Brightness and Saturation
        private void UpdateBrightnessBar(HSLColor color)
        {
            DrawBrightnessGradient(color);
        }

        private void BrightnessBar_ValueChanged(object sender, EventArgs e)
        {
            if (!timer1.Enabled)
            {
                float newBrightness = brightnessBar.Value;
                var newColor = new HSLColor();
                try
                {
                    newColor.Hue = dataSource.SelectedColor.Hue;
                    newColor.Saturation = dataSource.SelectedColor.Saturation;
                    newColor.Luminosity = newBrightness;

                    brightnessBar.ValueChanged -= BrightnessBar_ValueChanged;
                    saturationBar.ValueChanged -= SaturationBar_ValueChanged;

                    dataSource.SelectedColor = newColor;
                    UpdateSaturationBar(newColor);
                    UpdateBrightnessBar(newColor);

                    brightnessBar.ValueChanged += BrightnessBar_ValueChanged;
                    saturationBar.ValueChanged += SaturationBar_ValueChanged;
                }
                catch (Exception)
                {
                }
            }
        }

        private void SaturationBar_ValueChanged(object sender, EventArgs e)
        {
            float newSaturation = saturationBar.Value;
            var newColor = new HSLColor();
            try
            {
                newColor.Hue = dataSource.SelectedColor.Hue;
                newColor.Saturation = newSaturation;
                newColor.Luminosity = dataSource.SelectedColor.Luminosity;

                brightnessBar.ValueChanged -= BrightnessBar_ValueChanged;
                saturationBar.ValueChanged -= SaturationBar_ValueChanged;

                dataSource.SelectedColor = newColor;
                UpdateBrightnessBar(newColor);
                UpdateSaturationBar(newColor);
                
                brightnessBar.ValueChanged += BrightnessBar_ValueChanged;
                saturationBar.ValueChanged += SaturationBar_ValueChanged;
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
            Color[] blendColors = 
            {
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
                240)
            };
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
            Color[] blendColors = 
            {
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
                color.Luminosity)
            };
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
            DrawBrightnessGradient(dataSource.SelectedColor);
            DrawSaturationGradient(dataSource.SelectedColor);
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
            }
            else if (sender is Panel)
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

        private void Panel_DragDrop(object sender, DragEventArgs e)
        {
            Panel panel = (Panel)sender;
            panel.BackColor = (Color)e.Data.GetData(panel.BackColor.GetType());
            if (panel.Equals(panel1))
            {
                dataSource.SelectedColor = panel.BackColor;
            }
            if (panel.Equals(themePanel1))
            {
                dataSource.ThemeColorOne = panel.BackColor;
            }
            if (panel.Equals(themePanel2))
            {
                dataSource.ThemeColorTwo = panel.BackColor;
            }
            if (panel.Equals(themePanel3))
            {
                dataSource.ThemeColorThree = panel.BackColor;
            }
            if (panel.Equals(themePanel4))
            {
                dataSource.ThemeColorFour = panel.BackColor;
            }
            if (panel.Equals(themePanel5))
            {
                dataSource.ThemeColorFive = panel.BackColor;
            }
            if (panel.Equals(themePanel6))
            {
                dataSource.ThemeColorSix = panel.BackColor;
            }
            if (panel.Equals(themePanel7))
            {
                dataSource.ThemeColorSeven = panel.BackColor;
            }
            if (panel.Equals(themePanel8))
            {
                dataSource.ThemeColorEight = panel.BackColor;
            }
            if (panel.Equals(themePanel9))
            {
                dataSource.ThemeColorNine = panel.BackColor;
            }
            if (panel.Equals(themePanel10))
            {
                dataSource.ThemeColorTen = panel.BackColor;
            }
            UpdateUIForNewColor();
        }

        private void Panel_DragEnter(object sender, DragEventArgs e)
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
            dataSource.SelectedColor = clickedColor;
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
                    dataSource.ThemeColorOne = Color.White;
                    dataSource.ThemeColorTwo = Color.White;
                    dataSource.ThemeColorThree = Color.White;
                    dataSource.ThemeColorFour = Color.White;
                    dataSource.ThemeColorFive = Color.White;
                    dataSource.ThemeColorSix = Color.White;
                    dataSource.ThemeColorSeven = Color.White;
                    dataSource.ThemeColorEight = Color.White;
                    dataSource.ThemeColorNine = Color.White;
                    dataSource.ThemeColorTen = Color.White;
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
        private void ShowMoreInformationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)(contextMenuStrip1.SourceControl)).BackColor;
            ColorInformationDialog dialog = new ColorInformationDialog(clickedColor);
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.Show();
        }

        private void SelectAsMainColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)contextMenuStrip1.SourceControl).BackColor;

            dataSource.SelectedColor = clickedColor;
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            UpdateUIForNewColor();
        }

        private void AddToFavoritesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)contextMenuStrip1.SourceControl).BackColor;

            dataSource.ThemeColorTen = dataSource.ThemeColorNine;
            dataSource.ThemeColorNine = dataSource.ThemeColorEight;
            dataSource.ThemeColorEight = dataSource.ThemeColorSeven;
            dataSource.ThemeColorSeven = dataSource.ThemeColorSix;
            dataSource.ThemeColorSix = dataSource.ThemeColorFive;
            dataSource.ThemeColorFive = dataSource.ThemeColorFour;
            dataSource.ThemeColorFour = dataSource.ThemeColorThree;
            dataSource.ThemeColorThree = dataSource.ThemeColorTwo;
            dataSource.ThemeColorTwo = dataSource.ThemeColorOne;
            dataSource.ThemeColorOne = clickedColor;
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
                return dataSource.SelectedColor;

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

            return dataSource.SelectedColor;
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
                else if (!colorToReturn.Equals(color))
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
                case MODE.LINE:
                    return Color.FromArgb(ColorHelper.ReverseRGBToArgb(selectedShape.Line.ForeColor.RGB));
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
            return dataSource.SelectedColor;
        }

        private void SetModeForSenderName(string buttonName)
        {
            switch (buttonName)
            {
                case "fillButton":
                    SetMode(MODE.FILL);
                    break;
                case "fontButton":
                    SetMode(MODE.FONT);
                    break;
                case "lineButton":
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

        # region Functional Test APIs

        public Panel GetDropletPanel()
        {
            return panel1;
        }

        public Button GetFontColorButton()
        {
            return fontButton;
        }

        public Button GetLineColorButton()
        {
            return lineButton;
        }

        public Button GetFillCollorButton()
        {
            return fillButton;
        }

        public Panel GetMonoPanel1()
        {
            return monoPanel1;
        }

        public Panel GetMonoPanel7()
        {
            return monoPanel7;
        }

        public Panel GetFavColorPanel1()
        {
            return themePanel1;
        }

        public Button GetResetFavColorsButton()
        {
            return resetThemeButton;
        }

        public Button GetEmptyFavColorsButton()
        {
            return emptyPanelButton;
        }

        public ColorInformationDialog ShowMoreColorInfo(Color color)
        {
            ColorInformationDialog dialog = new ColorInformationDialog(color);
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.Show();
            return dialog;
        }

        # endregion
    }
}
