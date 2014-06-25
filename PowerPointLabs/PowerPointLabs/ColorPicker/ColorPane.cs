using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using System.Drawing.Drawing2D;
using PPExtraEventHelper;
using Converters = PowerPointLabs.Converters;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using PowerPointLabs.ColorPicker;
using PowerPointLabs.Views;

namespace PowerPointLabs
{

    public partial class ColorPane : UserControl
    {

        private Color _originalColor;

        //Defaults is Fill
        private bool _isFillColorSelected = true;
        private bool _isFontColorSelected = false;
        private bool _isLineColorSelected = false;

        private System.Drawing.Point _mouseDownLocation;
        LMouseListener _native;

        PowerPoint.ShapeRange _selectedShapes;
        ColorDataSource dataSource = new ColorDataSource();
        public ColorPane()
        {
            InitializeComponent();

            BindDataToPanels();

            InitToolTipControl();

            ResetThemePanel();
        }

        #region ToolTip
        private void InitToolTipControl()
        {
            toolTip1.SetToolTip(this.FontEyeDropperButton, "EyeDrops Font Color for Selected TextFrames");
            toolTip1.SetToolTip(this.LineEyeDropperButton, "EyeDrops Line Color for Selected Shapes");
            toolTip1.SetToolTip(this.FillEyeDropperButton, "EyeDrops Fill Color for Selected Shapes");
            toolTip1.SetToolTip(this.EditColorButton, "Edits Selected Color");
            toolTip1.SetToolTip(this.LoadButton, "Load Existing Theme");
            toolTip1.SetToolTip(this.SaveThemeButton, "Save Current Theme");
            toolTip1.SetToolTip(this.ResetThemeButton, "Reset the Current Theme");
        }
        #endregion

        #region DataBindings
        private void BindDataToPanels()
        {
            this.panel1.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "selectedColor",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel1.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorOne",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel2.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorTwo",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel3.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorThree",
                false,
                DataSourceUpdateMode.OnPropertyChanged));
            
            this.ThemePanel4.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorFour",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel5.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorFive",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel6.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorSix",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel7.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorSeven",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel8.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorEight",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel9.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorNine",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.ThemePanel10.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "themeColorTen",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

            this.AnalogousSelected.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "selectedColor",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

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

            this.TriadicSelected.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "selectedColor",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

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

            this.TetradicSelected.DataBindings.Add(new Binding(
                "BackColor",
                dataSource,
                "selectedColor",
                false,
                DataSourceUpdateMode.OnPropertyChanged));

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
        }

        #endregion

        private void BeginEyedropping()
        {
            timer1.Start();

            _native = new LMouseListener();
            _native.LButtonClicked +=
                 new EventHandler<SysMouseEventInfo>(_native_LButtonClicked);
            SelectShapes();
        }

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

        private void button2_Click(object sender, EventArgs e)
        {
            colorDialog1.Color = panel1.BackColor;
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                _originalColor = colorDialog1.Color;
                dataSource.selectedColor = colorDialog1.Color;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.Cross;
            System.Drawing.Point p = Control.MousePosition;
            IntPtr dc = Native.GetDC(IntPtr.Zero);
            dataSource.selectedColor = ColorTranslator.FromWin32(Native.GetPixel(dc, p.X, p.Y));
            ColorSelectedShapesWithColor(panel1.BackColor);
        }

        private void ColorSelectedShapesWithColor(Color selectedColor)
        {
            SelectShapes();
            if (_selectedShapes != null)
            {
                foreach (PowerPoint.Shape s in _selectedShapes)
                {
                    try
                    {
                        var r = selectedColor.R;
                        var g = selectedColor.G;
                        var b = selectedColor.B;

                        var rgb = (b << 16) | (g << 8) | (r);
                        ColorShapeWithColor(s, rgb);
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
            Shape newShape = PowerPointPresentation.CurrentSlide.Shapes.Paste()[1];

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

        private void ColorShapeWithColor(PowerPoint.Shape s, int rgb)
        {
            if (_isFillColorSelected)
            {
                s.Fill.ForeColor.RGB = rgb;
            }
            if (_isLineColorSelected)
            {
                s.Line.ForeColor.RGB = rgb;
            }
            if (_isFontColorSelected)
            {
                if (s.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    s.TextFrame.TextRange.Font.Color.RGB = rgb;
                }
            }
        }

        void _native_LButtonClicked(object sender, SysMouseEventInfo e)
        {
            _native.Close();
            _originalColor = panel1.BackColor;
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            UpdateUIForNewColor();
            timer1.Stop();
        }

        private void UpdateUIForNewColor()
        {
            UpdateBrightnessBar(dataSource.selectedColor);
            UpdateSaturationBar(dataSource.selectedColor);
            ColorSelectedShapesWithColor(dataSource.selectedColor);
        }

        private void UpdateBrightnessBar(Color color)
        {
            DrawBrightnessGradient(color);
        }

        private void UpdateSaturationBar(Color color)
        {
            DrawSaturationGradient(color);
        }

        protected override void OnPaint(PaintEventArgs paintEvnt)
        {
            DrawBrightnessGradient(dataSource.selectedColor);
            DrawSaturationGradient(dataSource.selectedColor);
        }

        private void DrawBrightnessGradient(Color color)
        {
            var dis = brightnessPanel.DisplayRectangle;
            var screenRec = brightnessPanel.RectangleToScreen(dis);
            var rec = brightnessPanel.Parent.RectangleToClient(screenRec);
            brightnessPanel.Visible = false;
            LinearGradientBrush brush = new LinearGradientBrush(
                rec,
                ColorHelper.ColorFromAhsb(
                255,
                color.GetHue(),
                color.GetSaturation(),
                0.00f),
                ColorHelper.ColorFromAhsb(
                255,
                color.GetHue(),
                color.GetSaturation(),
                1.0f),
                LinearGradientMode.Horizontal);
            ColorBlend blend = new ColorBlend();
            Color[] blendColors = {
                ColorHelper.ColorFromAhsb(
                255, 
                color.GetHue(),
                color.GetSaturation(),
                0.0f),
                color,
                ColorHelper.ColorFromAhsb(
                255, 
                color.GetHue(),
                color.GetSaturation(),
                1.0f)};
            float[] positions = { 0.0f, 0.5f, 1.0f };
            blend.Colors = blendColors;
            blend.Positions = positions;

            brush.InterpolationColors = blend;

            using (Graphics g = this.CreateGraphics())
            {
                g.FillRectangle(brush, rec);
            }
        }

        private void DrawSaturationGradient(Color color)
        {
            var dis = saturationPanel.DisplayRectangle;
            var screenRec = saturationPanel.RectangleToScreen(dis);
            var rec = saturationPanel.Parent.RectangleToClient(screenRec);
            saturationPanel.Visible = false;
            LinearGradientBrush brush = new LinearGradientBrush(
                rec,
                ColorHelper.ColorFromAhsb(
                255,
                color.GetHue(),
                0.0f,
                color.GetBrightness()),
                ColorHelper.ColorFromAhsb(
                255,
                color.GetHue(),
                1.0f,
                color.GetBrightness()),
                LinearGradientMode.Horizontal);
            ColorBlend blend = new ColorBlend();
            Color[] blendColors = {
                ColorHelper.ColorFromAhsb(
                255, 
                color.GetHue(),
                0.0f,
                color.GetBrightness()),
                color,
                ColorHelper.ColorFromAhsb(
                255, 
                color.GetHue(),
                1.0f,
                color.GetBrightness())};
            float[] positions = { 0.0f, 0.5f, 1.0f };
            blend.Colors = blendColors;
            blend.Positions = positions;

            brush.InterpolationColors = blend;

            using (Graphics g = this.CreateGraphics())
            {
                g.FillRectangle(brush, rec);
            }
        }

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
            _originalColor = panel.BackColor;
            UpdateUIForNewColor();
        }

        private void brightnessBar_ValueChanged(object sender, EventArgs e)
        {
            if (!timer1.Enabled)
            {
                float newBrightness = brightnessBar.Value / 240.0f;
                Color newColor = new Color();
                try
                {
                    newColor = ColorHelper.ColorFromAhsb(
                    255,
                    _originalColor.GetHue(),
                    dataSource.selectedColor.GetSaturation(),
                    newBrightness);

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
            float newSaturation = saturationBar.Value / 240.0f;
            Color newColor = new Color();
            try
            {
                newColor = ColorHelper.ColorFromAhsb(
                255,
                _originalColor.GetHue(),
                newSaturation,
                dataSource.selectedColor.GetBrightness());

                brightnessBar.ValueChanged -= brightnessBar_ValueChanged;
                saturationBar.ValueChanged -= saturationBar_ValueChanged;

                dataSource.selectedColor = newColor;
                UpdateBrightnessBar(newColor);

                brightnessBar.ValueChanged += brightnessBar_ValueChanged;
                saturationBar.ValueChanged += saturationBar_ValueChanged;
            }
            catch (Exception exception)
            {
            }

            ColorSelectedShapesWithColor(newColor);
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

        private void SaveThemeButton_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK && 
                dataSource.SaveThemeColorsInFile(saveFileDialog1.FileName))
            {
                MessageBox.Show("Theme saved successfully", "Save Complete");
            }
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK &&
                dataSource.LoadThemeColorsFromFile(openFileDialog1.FileName))
            {
                MessageBox.Show("Theme loaded successfully", "Load Complete");
            }
        }

        private void FontEyeDropperButton_Click(object sender, EventArgs e)
        {
            _isFontColorSelected = true;
            _isFillColorSelected = false;
            _isLineColorSelected = false;
            BeginEyedropping();
        }

        private void HighlightEyeDropperButton_Click(object sender, EventArgs e)
        {
            _isFontColorSelected = false;
            _isFillColorSelected = false;
            _isLineColorSelected = false;
            BeginEyedropping();
        }

        private void LineEyeDropperButton_Click(object sender, EventArgs e)
        {
            _isFontColorSelected = false;
            _isFillColorSelected = false;
            _isLineColorSelected = true;
            BeginEyedropping();
        }

        private void FillEyeDropperButton_Click(object sender, EventArgs e)
        {
            _isFontColorSelected = false;
            _isFillColorSelected = true;
            _isLineColorSelected = false;
            BeginEyedropping();
        }

        private void ThemePanel_Click(object sender, EventArgs e)
        {
            try 
	        {
                // Done twice due to multithreading issues with binding
                Color clickedColor = ((Panel)sender).BackColor;
                _originalColor = clickedColor;
                dataSource.selectedColor = clickedColor;
                UpdateUIForNewColor();

                _originalColor = clickedColor;
                dataSource.selectedColor = clickedColor;
                Globals.ThisAddIn.Application.StartNewUndoEntry();
                UpdateUIForNewColor();
	        }
	        catch (Exception ex)
	        {
                System.Diagnostics.Debug.WriteLine("Exception: " + ex.StackTrace);
		        throw;
	        }
        }

        private void MatchingPanel_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)sender).BackColor;
            
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            ColorSelectedShapesWithColor(clickedColor);
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

        private void FontEyeDropperButton_MouseDown(object sender, MouseEventArgs e)
        {
            _isFontColorSelected = true;
            _isFillColorSelected = false;
            _isLineColorSelected = false;
            BeginEyedropping();
        }

        private void LineEyeDropperButton_MouseDown(object sender, MouseEventArgs e)
        {
            _isFontColorSelected = false;
            _isFillColorSelected = false;
            _isLineColorSelected = true;
            BeginEyedropping();
        }

        private void FillEyeDropperButton_MouseDown(object sender, MouseEventArgs e)
        {
            _isFontColorSelected =false;
            _isFillColorSelected = true;
            _isLineColorSelected = false;
            BeginEyedropping();
        }

        private void ResetThemeButton_Click(object sender, EventArgs e)
        {
            ResetThemePanel();
        }

        private void ResetThemePanel()
        {
            try
            {
                Microsoft.Office.Core.ThemeColorScheme scheme = 
                    Globals.ThisAddIn.Application.ActivePresentation.Slides[1].ThemeColorScheme;
                
                ThemePanel1.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight1).RGB));
                ThemePanel2.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark1).RGB));
                ThemePanel3.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeLight2).RGB));
                ThemePanel4.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeDark2).RGB));
                ThemePanel5.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent1).RGB));
                ThemePanel6.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent2).RGB));
                ThemePanel7.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent3).RGB));
                ThemePanel8.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent4).RGB));
                ThemePanel9.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent5).RGB));
                ThemePanel10.BackColor = Color.FromArgb(
                    ColorHelper.ReverseRGBToArgb(scheme.Colors(Microsoft.Office.Core.MsoThemeColorSchemeIndex.msoThemeAccent6).RGB));
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Theme Panel Reset Failed", e.Message, e);
            }
        }

        private void ApplyCurrentThemeToSlides()
        {
            foreach (PowerPoint.Slide slide in Globals.ThisAddIn.Application.ActivePresentation.Slides)
            {
                ApplyCurrentThemeToSlide(slide);
            }
        }

        private void ApplyCurrentThemeToSlide(Slide slide)
        {
            Microsoft.Office.Core.ThemeColorScheme scheme = 
                slide.ThemeColorScheme;

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

        private void ColorPane_Load(object sender, EventArgs e)
        {
            ResetThemePanel();
        }

        private void showMoreInformationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)(contextMenuStrip1.SourceControl)).BackColor;
            ColorInformationDialog dialog = new ColorInformationDialog(clickedColor);
            dialog.StartPosition = FormStartPosition.CenterScreen;
            dialog.ShowDialog();
        }

        private void ApplyThemeButton_Click(object sender, EventArgs e)
        {
            ApplyCurrentThemeToSlides();
        }

        private void selectAsMainColorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Color clickedColor = ((Panel)contextMenuStrip1.SourceControl).BackColor;
            dataSource.selectedColor = clickedColor;
            _originalColor = clickedColor;
            Globals.ThisAddIn.Application.StartNewUndoEntry();
            UpdateUIForNewColor();
        }
    }

    public class SysMouseEventInfo : EventArgs
    {
        public string WindowTitle { get; set; }
    }
    public class LMouseListener
    {
        public LMouseListener()
        {
            this.CallBack += new HookProc(MouseEvents);
            //Module mod = Assembly.GetExecutingAssembly().GetModules()[0];
            //IntPtr hMod = Marshal.GetHINSTANCE(mod);
            using (System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess())
            using (System.Diagnostics.ProcessModule module = process.MainModule)
            {
                IntPtr hModule = GetModuleHandle(module.ModuleName);
                _hook = SetWindowsHookEx(WH_MOUSE_LL, this.CallBack, hModule, 0);
                //if (_hook != IntPtr.Zero)
                //{
                //    Console.WriteLine("Started");
                //}
            }
        }
        int WH_MOUSE_LL = 14;
        int HC_ACTION = 0;
        HookProc CallBack = null;
        IntPtr _hook = IntPtr.Zero;

        public event EventHandler<SysMouseEventInfo> LButtonClicked;

        int MouseEvents(int code, IntPtr wParam, IntPtr lParam)
        {
            //Console.WriteLine("Called");

            if (code < 0)
                return CallNextHookEx(_hook, code, wParam, lParam);

            if (code == this.HC_ACTION)
            {
                // Left button pressed somewhere
                if (wParam.ToInt32() == (uint)WM.WM_LBUTTONUP)
                {
                    MSLLHOOKSTRUCT ms = new MSLLHOOKSTRUCT();
                    ms = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    IntPtr win = WindowFromPoint(ms.pt);
                    string title = GetWindowTextRaw(win);
                    if (LButtonClicked != null)
                    {
                        LButtonClicked(this, new SysMouseEventInfo { WindowTitle = title });
                    }
                }
            }
            return CallNextHookEx(_hook, code, wParam, lParam);
        }

        public void Close()
        {
            if (_hook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hook);
            }
        }
        public delegate int HookProc(int code, IntPtr wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "SetWindowsHookEx", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        static extern IntPtr WindowFromPoint(int xPoint, int yPoint);

        [DllImport("user32.dll")]
        static extern IntPtr WindowFromPoint(POINT Point);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, [Out] StringBuilder lParam);

        public static string GetWindowTextRaw(IntPtr hwnd)
        {
            // Allocate correct string length first
            //int length = (int)SendMessage(hwnd, (int)WM.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);
            StringBuilder sb = new StringBuilder(65535);//THIS COULD BE BAD. Maybe you shoudl get the length
            SendMessage(hwnd, (int)WM.WM_GETTEXT, (IntPtr)sb.Capacity, sb);
            return sb.ToString();
        }
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public int mouseData;
        public int flags;
        public int time;
        public UIntPtr dwExtraInfo;
    }
    enum WM : uint
    {//all windows messages here
        WM_LBUTTONUP = 0x0202,
        WM_GETTEXT = 0x000D,
        WM_GETTEXTLENGTH = 0x000E
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;

        public POINT(int x, int y)
        {
            this.X = x;
            this.Y = y;
        }
    }
}
