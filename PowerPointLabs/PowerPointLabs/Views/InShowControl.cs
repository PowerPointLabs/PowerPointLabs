using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using Point = System.Drawing.Point;

namespace PowerPointLabs.Views
{
    public partial class InShowControl : Form
    {
        private enum ButtonStatus
        {
            Idle,
            Rec
        }

        private ButtonStatus _status;
        
        private int _currentSlide;
        private int _currentScriptIndex;
        private SlideShowWindow _slideShowWindow;

        public InShowControl()
        {
            InitializeComponent();
            
            // set the background transparency
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;

            _slideShowWindow = Globals.ThisAddIn.Application.ActivePresentation.SlideShowWindow;

            // get slide show window
            IntPtr slideShowWindow = new IntPtr(_slideShowWindow.HWND);
            
            Native.RECT rec;
            Native.GetWindowRect(new HandleRef(new object(), slideShowWindow), out rec);
            
            Location = new Point(rec.Right - Width, rec.Bottom - Height- 50);

            _currentScriptIndex = 0;
        }

        private void RecButtonClick(object sender, EventArgs e)
        {
            var recorderPane = Globals.ThisAddIn.recorderTaskPane;
            var click = _slideShowWindow.View.GetClickIndex();
            var currentSlide = PowerPointSlide.FromSlideFactory(_slideShowWindow.View.Slide);

            switch (_status)
            {
                case ButtonStatus.Idle:
                    _status = ButtonStatus.Rec;
                    recButton.Text = "Stop and Advance";
                    recorderPane.RecButtonIdleHandler();
                    break;

                case ButtonStatus.Rec:
                    _status = ButtonStatus.Idle;
                    recButton.Text = "Record";
                    
                    recorderPane.StopButtonRecordingHandler(recorderPane.GetPlaybackFromList(click, currentSlide.ID),
                                                            click, currentSlide);
                    _slideShowWindow.View.GotoClick(click + 1);
                    break;
            }
        }
    }
}
