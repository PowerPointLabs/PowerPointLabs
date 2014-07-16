using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
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
        
        // Data-bindings datasource
        ColorDataSource dataSource = new ColorDataSource();

        // Stores last selected mode
        private MODE prevMode = MODE.NONE;
        
        // Stores the current selected mode
        private MODE currMode = MODE.NONE;

        public ColorPane()
        {
            InitializeComponent();

            BindDataToPanels();

            InitToolTipControl();

            ResetThemePanel();

            // Default color to CornFlowerBlue
            SetDefaultColor(Color.CornflowerBlue);
        }
        private void ColorPane_Load(object sender, EventArgs e)
        {
            ResetThemePanel();
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
            toolTip1.SetToolTip(this.LoadButton, "Load Existing Theme");
            toolTip1.SetToolTip(this.SaveThemeButton, "Save Current Theme");
            toolTip1.SetToolTip(this.ResetThemeButton, "Reset the Current Theme Colors to Current Slide Theme");
            toolTip1.SetToolTip(this.FontButton, "Font Color\r\nMouseDown + Drag to EyeDrop\r\nClick to Edit");
            toolTip1.SetToolTip(this.LineButton, "Shape Line Color\r\nMouseDown + Drag to EyeDrop\r\nClick to Edit");
            toolTip1.SetToolTip(this.FillButton, "Shape Fill Color\r\nMouseDown + Drag to EyeDrop\r\nClick to Edit");
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

        private void BeginEyedropping()
        {
            timer1.Start();
            _native = new LMouseUpListener();
            _native.LButtonUpClicked +=
                 new EventHandler<SysMouseEventInfo>(_native_LButtonClicked);
            SelectShapes();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.Cross;
            System.Drawing.Point p = Control.MousePosition;
            IntPtr dc = Native.GetDC(IntPtr.Zero);
            dataSource.selectedColor = ColorTranslator.FromWin32(Native.GetPixel(dc, p.X, p.Y));
            ColorSelectedShapesWithColor(panel1.BackColor);
        }

        #region Selection And Coloring Shapes
        private void SelectShapes()
        {
            try
            {
                _selectedShapes = PowerPointPresentation.CurrentSelection.ShapeRange;
            }
            catch (Exception exception)
            {
                _selectedShapes = null;
            }
        }
        private void ColorSelectedShapesWithColor(HSLColor selectedColor)
        {
            SelectShapes();
            if (_selectedShapes != null)
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
                    catch (Exception e)
                    {
                        RecreateCorruptedShape(s);
                    }  
                }
            }
        }

        private static void RecreateCorruptedShape(PowerPoint.Shape s)
        {
            s.Copy();
            PowerPoint.Shape newShape = PowerPointPresentation.CurrentSlide.Shapes.Paste()[1];

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
                    break;
                case MODE.FONT:
                    if (s.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.HasTextFrame
                            == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            TextRange selectedText
                                = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.TrimText();
                            if (selectedText.Text != "" && selectedText != null)
                            {
                                selectedText.Font.Color.RGB = rgb;
                            }
                        }
                    }
                    break;
            }
        }

        void _native_LButtonClicked(object sender, SysMouseEventInfo e)
        {
            _native.Close();
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            UpdateUIForNewColor();
            timer1.Stop();
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
            ColorSelectedShapesWithColor(dataSource.selectedColor);
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
                catch (Exception exception)
                {
                }

                ColorSelectedShapesWithColor(dataSource.selectedColor);
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
            catch (Exception exception)
            {
            }

            ColorSelectedShapesWithColor(newColor);
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
                new HSLColor(
                color.Hue,
                color.Saturation,
                0),
                new HSLColor(
                color.Hue,
                color.Saturation,
                240),
                LinearGradientMode.Horizontal);
            ColorBlend blend = new ColorBlend();
            Color[] blendColors = {
                new HSLColor(
                color.Hue,
                color.Saturation,
                0),
                color,
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
                new HSLColor(
                color.Hue,
                0,
                color.Luminosity),
                new HSLColor(
                color.Hue,
                240,
                color.Luminosity),
                LinearGradientMode.Horizontal);
            ColorBlend blend = new ColorBlend();
            Color[] blendColors = {
                new HSLColor(
                color.Hue,
                0,
                color.Luminosity),
                color,
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
        private void MatchingPanel_MouseDown(object sender, MouseEventArgs e)
        {
            _mouseDownLocation = e.Location;
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
            UpdateUIForNewColor();
        }

        private void panel_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(dataSource.selectedColor.GetType().ToString()))
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
        private void SaveThemeButton_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK && 
                dataSource.SaveThemeColorsInFile(saveFileDialog1.FileName))
            {
                // Save Success
            }
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK &&
                dataSource.LoadThemeColorsFromFile(openFileDialog1.FileName))
            {
                // Load Success
            }
        }

        private void MatchingPanel_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)sender).BackColor;
            
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ColorSelectedShapesWithColor(clickedColor);
        }

        private void ResetThemeButton_Click(object sender, EventArgs e)
        {
            ResetThemePanel();
        }

        private void ApplyThemeButton_Click(object sender, EventArgs e)
        {
            ApplyCurrentThemeToSelectedSlides();
        }

        private void ResetThemePanel()
        {
            try
            {
                if (PowerPointPresentation.SlideCount > 0)
                {
                    Microsoft.Office.Core.ThemeColorScheme scheme =
                    PowerPointPresentation.CurrentSlide.GetNativeSlide().ThemeColorScheme;

                    ThemePanel1.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight1).RGB));
                    dataSource.themeColorOne = ThemePanel1.BackColor;

                    ThemePanel2.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark1).RGB));
                    dataSource.themeColorTwo = ThemePanel2.BackColor;

                    ThemePanel3.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight2).RGB));
                    dataSource.themeColorThree = ThemePanel3.BackColor;

                    ThemePanel4.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark2).RGB));
                    dataSource.themeColorFour = ThemePanel4.BackColor;

                    ThemePanel5.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent1).RGB));
                    dataSource.themeColorFive = ThemePanel5.BackColor;

                    ThemePanel6.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent2).RGB));
                    dataSource.themeColorSix = ThemePanel6.BackColor;

                    ThemePanel7.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent3).RGB));
                    dataSource.themeColorSeven = ThemePanel7.BackColor;

                    ThemePanel8.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent4).RGB));
                    dataSource.themeColorEight = ThemePanel8.BackColor;

                    ThemePanel9.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent5).RGB));
                    dataSource.themeColorNine = ThemePanel9.BackColor;

                    ThemePanel10.BackColor = Color.FromArgb(
                        ColorHelper.ReverseRGBToArgb(
                        scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent6).RGB));
                    dataSource.themeColorTen = ThemePanel10.BackColor;
                }
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
                if (PowerPointPresentation.SlideCount > 0)
                {
                    ThemePanel1.BackColor = Color.White;
                    ThemePanel2.BackColor = Color.White;
                    ThemePanel3.BackColor = Color.White;
                    ThemePanel4.BackColor = Color.White;
                    ThemePanel5.BackColor = Color.White;
                    ThemePanel6.BackColor = Color.White;
                    ThemePanel7.BackColor = Color.White;
                    ThemePanel8.BackColor = Color.White;
                    ThemePanel9.BackColor = Color.White;
                    ThemePanel10.BackColor = Color.White;
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Theme Panel Reset Failed", e.Message, e);
            }
        }

        private void ApplyCurrentThemeToSelectedSlides()
        {
            foreach (PowerPointSlide slide in PowerPointPresentation.SelectedSlides)
            {
                ApplyCurrentThemeToSlide(slide);
            }
        }

        private void ApplyCurrentThemeToSlide(PowerPointSlide slide)
        {
            Microsoft.Office.Core.ThemeColorScheme scheme = 
                slide.GetNativeSlide().ThemeColorScheme;

            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight1).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel1.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark1).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel2.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight2).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel3.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark2).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel4.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent1).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel5.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent2).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel6.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent3).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel7.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent4).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel8.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent5).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel9.BackColor.ToArgb()));
            scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent6).RGB =
                ColorHelper.ReverseRGBToArgb((ThemePanel10.BackColor.ToArgb()));
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
            // Done twice due to multi-threading issues

            Color clickedColor = ((Panel)contextMenuStrip1.SourceControl).BackColor;
            dataSource.selectedColor = clickedColor;
            //Globals.ThisAddIn.Application.StartNewUndoEntry();
            UpdateUIForNewColor();

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

            colorDialog1.Color = dataSource.selectedColor;

            DialogResult result = colorDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                SetDefaultColor(colorDialog1.Color);
                
                ResetEyeDropperSelectionInDataSource();
            }
            else if (result == DialogResult.Cancel)
            {
                ResetEyeDropperSelectionInDataSource();
            }
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
    }
}
