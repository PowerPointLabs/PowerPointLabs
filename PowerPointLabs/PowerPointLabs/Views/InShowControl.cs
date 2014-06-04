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
        public enum ButtonStatus
        {
            Idle,
            Rec
        }

        private ButtonStatus _status;
        private SlideShowWindow _slideShowWindow;
        private int _startingSlide;
        private int _recordStartClick;
        private PowerPointSlide _recordStartSlide;

        public InShowControl()
        {
            InitializeComponent();
            
            // set the background transparency
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;

            _slideShowWindow = Globals.ThisAddIn.Application.ActivePresentation.SlideShowWindow;
            _startingSlide = Globals.ThisAddIn.Application.ActivePresentation.SlideShowSettings.StartingSlide;

            // get slide show window
            IntPtr slideShowWindow = new IntPtr(_slideShowWindow.HWND);
            
            Native.RECT rec;
            Native.GetWindowRect(new HandleRef(new object(), slideShowWindow), out rec);
            
            Location = new Point(rec.Right - Width, rec.Bottom - Height- 55);
            _status = ButtonStatus.Idle;
        }

        public ButtonStatus GetCurrentStatus()
        {
            return _status;
        }

        public void ForceStop()
        {
            _status = ButtonStatus.Idle;
            Globals.ThisAddIn.recorderTaskPane.StopButtonRecordingHandler(_recordStartClick, _recordStartSlide, false);
        }

        private int GetRelativeSlideIndex(int index)
        {
            return index - _startingSlide + 1;
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
                    _recordStartClick = click;
                    _recordStartSlide = currentSlide;

                    recButton.Text = "Stop and Advance";
                    recorderPane.RecButtonIdleHandler();
                    _slideShowWindow.Activate();
                    break;

                case ButtonStatus.Rec:
                    recButton.Text = "Start Recording";

                    recorderPane.StopButtonRecordingHandler(_recordStartClick, _recordStartSlide, true);
                    _slideShowWindow.Activate();

                    var totalClick = _slideShowWindow.View.GetClickCount();

                    if (click + 1 > totalClick)
                    {
                        _slideShowWindow.View.Next();
                    }
                    else
                    {
                        _slideShowWindow.View.GotoClick(click + 1);
                    }

                    _status = ButtonStatus.Idle;
                    
                    break;
            }
        }

        private void UndoButtonClick(object sender, EventArgs e)
        {
            var recorderPane = Globals.ThisAddIn.recorderTaskPane;
            var temp = recorderPane.AudioBuffer[_recordStartSlide.Index - 1];

            // revert back the last action
            recorderPane.UndoLastRecord(_recordStartClick, _recordStartSlide);
            temp.RemoveAt(temp.Count - 1);

            // goto previous slide and click
            _slideShowWindow.View.GotoSlide(GetRelativeSlideIndex(_recordStartSlide.Index));
            _slideShowWindow.View.GotoClick(_recordStartClick);

            // activate the show window
            _slideShowWindow.Activate();
        }
    }
}
