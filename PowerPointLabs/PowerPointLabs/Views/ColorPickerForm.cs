using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using PowerPointLabs.Models;

namespace PowerPointLabs.Views
{
    public partial class ColorPickerForm : Form
    {
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

        class Win32API
        {
            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool GetCursorPos(out POINT lpPoint);
        }

        private POINT GetCursorPosition()
        {
            POINT point = new POINT();

            Win32API.GetCursorPos(out point);

            Win32API.ScreenToClient(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref point);

            return point;
        }
        private POINT ConvertScreenPointToSlideCoordinates(POINT point)
        {
            // Get the screen coordinates of the upper-left hand corner of the slide.
            POINT refPointUpperLeft = new POINT(
                Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(0),
                Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(0));

            // Get the size of the slide (in points of the slide's coordinate system).
            var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            var slideWidth = slide.CustomLayout.Width;
            var slideHeight = slide.CustomLayout.Height;

            // Get the screen coordinates of the bottom-right hand corner of the slide.
            POINT refPointBottomRight = new POINT(
                Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsX(slideWidth),
                Globals.ThisAddIn.Application.ActiveWindow.PointsToScreenPixelsY(slideHeight));

            // Both reference points have to be converted to the PowerPoint window's coordinate system.
            Win32API.ScreenToClient(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref refPointUpperLeft);
            Win32API.ScreenToClient(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND), ref refPointBottomRight);

            // Convert the point to the slide's coordinate system (convert the pixel coordinate inside the slide into a 0..1 interval, then scale it up by the slide's point size).
            return new POINT(
                (int)Math.Round((double)(point.X - refPointUpperLeft.X) / (refPointBottomRight.X - refPointUpperLeft.X) * slideWidth),
                (int)Math.Round((double)(point.Y - refPointUpperLeft.Y) / (refPointBottomRight.Y - refPointUpperLeft.Y) * slideHeight));
        }
        // Refactor this to Native

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

        public ColorPickerForm()
        {
            InitializeComponent();
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            this.timer1.Start();
            //start
            _native = new LMouseListener();
            _native.LButtonClicked +=
                 new EventHandler<SysMouseEventInfo>(_native_LButtonClicked);

        }

        LMouseListener _native;

        void _native_LButtonClicked(object sender, SysMouseEventInfo e)
        {
            this.Activate();
            timer1.Stop();
            _native.Close();
        }
        public ColorPickerForm(PowerPoint.ShapeRange selectedShapes)
            : this()
        {

        }

        private void ColorPickerForm_Load(object sender, EventArgs e)
        {

        }


        private void ColorPickerForm_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
        {
            Random r = new Random(1);

            for (int x = 0; x < 100; x++)
            {
                SolidBrush b = new SolidBrush(Color.FromArgb(r.Next(255), r.Next(255), r.Next(255)));
                e.Graphics.FillRectangle(b, r.Next(this.ClientSize.Width), r.Next(this.ClientSize.Height), r.Next(100), r.Next(100));
            }
        }

        private void timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Point p = Control.MousePosition;
            IntPtr dc = GetDC(IntPtr.Zero);
            this.panel1.BackColor = ColorTranslator.FromWin32(GetPixel(dc, p.X, p.Y));

        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = Control.MousePosition;
            IntPtr dc = GetDC(IntPtr.Zero);
            this.panel1.BackColor = ColorTranslator.FromWin32(GetPixel(dc, p.X, p.Y));
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            colorDialog1.Color = panel1.BackColor;
            colorDialog1.FullOpen = true;
            colorDialog1.ShowDialog();
            
        }

        private void panel1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Hello");
            _native = new LMouseListener();
            _native.LButtonClicked +=
                 new EventHandler<SysMouseEventInfo>(_native_LButtonClicked2);
        }
        void _native_LButtonClicked2(object sender, SysMouseEventInfo e)
        {
            var point = GetCursorPosition();
            var convertedPoint = this.ConvertScreenPointToSlideCoordinates(point);

            // Insert something at the cursor's location:
            var slide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            slide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeCloud, convertedPoint.X, convertedPoint.Y, 100, 100);
            _native.Close();
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
