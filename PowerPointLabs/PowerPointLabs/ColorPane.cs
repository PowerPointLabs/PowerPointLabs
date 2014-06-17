using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;

namespace PowerPointLabs
{
    public partial class ColorPane : UserControl
    {
        private Color _selectedColor;
        [DllImport("Gdi32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int GetPixel(
            System.IntPtr hdc,    // handle to DC
            int nXPos,  // x-coordinate of pixel
            int nYPos   // y-coordinate of pixel
        );

        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GetDC(IntPtr wnd);

        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern void ReleaseDC(IntPtr dc);

        LMouseListener _native;

        PowerPoint.ShapeRange _selectedShapes;
        public ColorPane()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            timer1.Start();

            _native = new LMouseListener();
            _native.LButtonClicked +=
                 new EventHandler<SysMouseEventInfo>(_native_LButtonClicked);
            //DisableMouseClicks();

            SelectShapes();
        }

        private void SelectShapes()
        {
            try
            {
                _selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            }
            catch (Exception exception)
            {
                _selectedShapes = null;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            colorDialog1.Color = panel1.BackColor;
            colorDialog1.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = Control.MousePosition;
            IntPtr dc = GetDC(IntPtr.Zero);
            this.panel1.BackColor = ColorTranslator.FromWin32(GetPixel(dc, p.X, p.Y));
            UpdatePanelsForColor(panel1.BackColor);
            ColorSelectedShapesWithColor(panel1.BackColor);
        }

        private void ColorSelectedShapesWithColor(Color selectedColor)
        {
            SelectShapes();
            if (_selectedShapes != null)
            {
                foreach (PowerPoint.Shape s in _selectedShapes)
                {
                    var r = selectedColor.R;
                    var g = selectedColor.G;
                    var b = selectedColor.B;

                    var rgb = (b << 16) | (g << 8) | (r);
                    s.Fill.ForeColor.RGB = rgb;
                }
            }
        }

        private void MatchingColorPanel_DoubleClick(object sender, EventArgs e)
        {
            Color selectedColor = ((Panel)sender).BackColor;

            UpdatePanelsForColor(selectedColor);
            ColorSelectedShapesWithColor(selectedColor);
        }

        void _native_LButtonClicked(object sender, SysMouseEventInfo e)
        {
            _native.Close();
            _selectedColor = panel1.BackColor;
            brightnessBar.Value = (int)(_selectedColor.GetBrightness() * 240.0f);
            UpdatePanelsForColor(_selectedColor);
            ColorSelectedShapesWithColor(_selectedColor);
            timer1.Stop();
            //EnableMouseClicks();
        }

        private void GenerateButton_Click(object sender, EventArgs e)
        {
            UpdatePanelsForColor(panel1.BackColor);
        }

        private void UpdatePanelsForColor(Color selectedColor)
        {
            panel1.BackColor = selectedColor;
            Color complementaryColor = ColorHelper.GetComplementaryColor(selectedColor);

            List<Color> analogousColors = ColorHelper.GetAnalogousColorsForColor(selectedColor);
            AnalogousLighter.BackColor = analogousColors[0];
            AnalogousDarker.BackColor = analogousColors[1];
            AnalogousSelected.BackColor = selectedColor;

            List<Color> complementaryColors = ColorHelper.GetSplitComplementaryColorsForColor(selectedColor);
            ComplementaryLighter.BackColor = complementaryColors[0];
            ComplementaryDarker.BackColor = complementaryColors[1];
            ComplementarySelected.BackColor = complementaryColor;

            List<Color> triadicColors = ColorHelper.GetTriadicColorsForColor(selectedColor);
            TriadicLower.BackColor = triadicColors[0];
            TriadicHigher.BackColor = triadicColors[1];
            TriadicSelected.BackColor = selectedColor;

            List<Color> tetradicColors = ColorHelper.GetTetradicColorsForColor(selectedColor);
            Tetradic1.BackColor = tetradicColors[0];
            Tetradic2.BackColor = tetradicColors[1];
            Tetradic3.BackColor = tetradicColors[2];
            TetradicSelected.BackColor = selectedColor;
        }

        private void AnalogousLighter_MouseDown(object sender, MouseEventArgs e)
        {
            DataObject colorObject = new DataObject();
            colorObject.SetData(AnalogousLighter.BackColor);
            DoDragDrop(colorObject, DragDropEffects.All);
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            panel1.BackColor = (Color)e.Data.GetData(panel1.BackColor.GetType());
        }

        private void DisableMouseClicks()
        {
            if (this.Filter == null)
            {
                this.Filter = new MouseClickMessageFilter();
                Application.AddMessageFilter(this.Filter);
            }
        }

        private void EnableMouseClicks()
        {
            if ((this.Filter != null))
            {
                Application.RemoveMessageFilter(this.Filter);
                this.Filter = null;
            }
        }

        private MouseClickMessageFilter Filter;

        private const int LButtonDown = 0x0201;

        public class MouseClickMessageFilter : IMessageFilter
        {

            public bool PreFilterMessage(ref System.Windows.Forms.Message m)
            {
                switch (m.Msg)
                {
                    case LButtonDown:
                        return true;
                }
                return false;
            }
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
                    _selectedColor.GetHue(),
                    _selectedColor.GetSaturation(),
                    newBrightness);
                } 
                catch(Exception exception)
                {
                    System.Diagnostics.Debug.WriteLine(exception.StackTrace);
                }
                
                compareHue(newColor, _selectedColor);
                compareSaturation(newColor, _selectedColor);

                UpdatePanelsForColor(newColor);
                ColorSelectedShapesWithColor(newColor);
            }
        }

        private void compareColor(Color one, Color two)
        {
            compareHue(one, two);
            compareSaturation(one, two);
            compareBrightness(one, two);
        }

        private static void compareBrightness(Color one, Color two)
        {
            if (one.GetBrightness() != two.GetBrightness())
            {
                System.Diagnostics.Debug.WriteLine("Brightness mismatch: " +
                    one.GetBrightness() + "\t" + two.GetBrightness());
            }
        }

        private static void compareSaturation(Color one, Color two)
        {
            if (one.GetSaturation() != two.GetSaturation())
            {
                System.Diagnostics.Debug.WriteLine("Saturation mismatch: " +
                    one.GetSaturation() + "\t" + two.GetSaturation());
            }
        }

        private static void compareHue(Color one, Color two)
        {
            if (one.GetHue() != two.GetHue())
            {
                System.Diagnostics.Debug.WriteLine("Hue mismatch: " +
                    one.GetHue() + "\t" + two.GetHue());
            }
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
