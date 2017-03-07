using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PPExtraEventHelper;

using Point = System.Drawing.Point;

namespace PowerPointLabs.Views
{
    internal partial class InShowControl : Form
    {
#pragma warning disable 0618
        public enum ButtonStatus
        {
            Idle,
            Estop,
            Rec
        }

        private ButtonStatus _status;
        private readonly SlideShowWindow _slideShowWindow;
        private readonly int _startingSlide;
        private readonly RecorderTaskPane _recorder;
        private int _recordStartClick;
        private PowerPointSlide _recordStartSlide;

        public InShowControl(RecorderTaskPane recorder)
        {
            InitializeComponent();
            
            // set the background transparency
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;

            _slideShowWindow = PowerPointPresentation.Current.Presentation.SlideShowWindow;
            _startingSlide = PowerPointPresentation.Current.Presentation.SlideShowSettings.StartingSlide;
            _recorder = recorder;

            // get slide show window
            var slideShowWindow = new IntPtr(_slideShowWindow.HWND);
            
            Native.RECT rec;
            Native.GetWindowRect(new HandleRef(new object(), slideShowWindow), out rec);
            
            Location = new Point(rec.Right - Width, rec.Bottom - Height- 65);
            _status = ButtonStatus.Idle;

            // disable undo button by default, enable only when there has something to undo
            undoButton.Enabled = false;
        }

        public ButtonStatus GetCurrentStatus()
        {
            return _status;
        }

        public void ForceStop()
        {
            if (_recorder != null)
            {
                _status = ButtonStatus.Estop;
                _recorder.StopButtonRecordingHandler(_recordStartClick, _recordStartSlide, false);
                _status = ButtonStatus.Idle;
            }
        }

        private int GetRelativeSlideIndex(int index)
        {
            return index - _startingSlide + 1;
        }

        private void RecButtonClick(object sender, EventArgs e)
        {
            if (_recorder == null) return;
            
            int click;
            PowerPointSlide currentSlide;

            try
            {
                click = _slideShowWindow.View.GetClickIndex();
                currentSlide = PowerPointSlide.FromSlideFactory(_slideShowWindow.View.Slide);
            }
            catch (COMException)
            {
                MessageBox.Show(TextCollection.InShowControlInvalidRecCommandError);
                return;
            }

            switch (_status)
            {
                case ButtonStatus.Idle:
                    _status = ButtonStatus.Rec;
                    undoButton.Enabled = false;
                    _recordStartClick = click;
                    _recordStartSlide = currentSlide;

                    recButton.Text = TextCollection.InShowControlRecButtonIdleText;
                    recButton.ForeColor = Color.Red;
                    _recorder.RecButtonIdleHandler();
                    _slideShowWindow.Activate();
                    break;

                case ButtonStatus.Rec:
                    recButton.Text = TextCollection.InShowControlRecButtonRecText;
                    undoButton.Enabled = true;
                    recButton.ForeColor = Color.Black;

                    _recorder.StopButtonRecordingHandler(_recordStartClick, _recordStartSlide, true);
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
                    
                    // stop produces a undo-able record, thus enable undo button
                    undoButton.Enabled = true;
                    
                    break;
            }
        }

        private void UndoButtonClick(object sender, EventArgs e)
        {
            if (_recorder == null) return;

            var temp = _recorder.AudioBuffer[_recordStartSlide.Index - 1];

            // disable undo since we allow only 1 step undo
            undoButton.Enabled = false;

            // revert back the last action
            _recorder.UndoLastRecord(_recordStartClick, _recordStartSlide);
            temp.RemoveAt(temp.Count - 1);

            // goto previous slide and click
            _slideShowWindow.View.GotoSlide(GetRelativeSlideIndex(_recordStartSlide.Index));
            _slideShowWindow.View.GotoClick(_recordStartClick);

            // activate the show window to allow user click on the slide show
            _slideShowWindow.Activate();
        }

        private bool InRectangle(Point point, Rectangle rect)
        {
            return point.X <= rect.Right && point.X >= rect.Left &&
                   point.Y >= rect.Top && point.Y <= rect.Bottom;
        }

        private void InShowControlMouseClick(object sender, MouseEventArgs e)
        {
            var control = sender as Control;

            if (control != null &&
                InRectangle(control.PointToScreen(e.Location),
                            undoButton.RectangleToScreen(undoButton.DisplayRectangle))
                && undoButton.Enabled == false)
            {
                _slideShowWindow.Activate();
            }
        }

        private void InShowControlMouseDoubleClick(object sender, MouseEventArgs e)
        {
            var control = sender as Control;

            if (control != null &&
                InRectangle(control.PointToScreen(e.Location),
                            undoButton.RectangleToScreen(undoButton.DisplayRectangle)) &&
                undoButton.Enabled == false)
            {
                _slideShowWindow.Activate();
            }
        }
    }
}
