﻿using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;

using PPExtraEventHelper;

using Forms = System.Windows.Forms;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for InShowRecordingControl.xaml
    /// </summary>
    public partial class InShowRecordingControl
    {
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

        public InShowRecordingControl()
        {
            InitializeComponent();
            AllowsTransparency = true;
        }

        public InShowRecordingControl(RecorderTaskPane recorder)
            : this()
        {
            _slideShowWindow = this.GetCurrentPresentation().Presentation.SlideShowWindow;
            _startingSlide = this.GetCurrentPresentation().Presentation.SlideShowSettings.StartingSlide;
            _recorder = recorder;

            // get slide show window
            var slideShowWindow = new IntPtr(_slideShowWindow.HWND);

            Native.RECT rec;
            Native.GetWindowRect(new HandleRef(new object(), slideShowWindow), out rec);
            
            Left = rec.Right / Graphics.GetDpiScale() - Width;
            Top = rec.Bottom / Graphics.GetDpiScale() - Height;
            _status = ButtonStatus.Idle;

            // disable undo button by default, enable only when there has something to undo
            undoButton.IsEnabled = false;
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

        private void RecButton_Click(object sender, RoutedEventArgs e)
        {
            if (_recorder == null)
            {
                return;
            }

            int click;
            PowerPointSlide currentSlide;

            try
            {
                click = _slideShowWindow.View.GetClickIndex();
                currentSlide = PowerPointSlide.FromSlideFactory(_slideShowWindow.View.Slide);
            }
            catch (COMException)
            {
                Forms.MessageBox.Show(TextCollection.InShowControlInvalidRecCommandError);
                return;
            }

            switch (_status)
            {
                case ButtonStatus.Idle:
                    _status = ButtonStatus.Rec;
                    undoButton.IsEnabled = false;
                    _recordStartClick = click;
                    _recordStartSlide = currentSlide;

                    recButton.Content = TextCollection.InShowControlRecButtonIdleText;
                    recButton.Foreground = new SolidColorBrush(Colors.Red);
                    _recorder.RecButtonIdleHandler();
                    _slideShowWindow.Activate();
                    break;

                case ButtonStatus.Rec:
                    recButton.Content = TextCollection.InShowControlRecButtonRecText;
                    undoButton.IsEnabled = true;
                    recButton.Foreground = new SolidColorBrush(Colors.Black);

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
                    undoButton.IsEnabled = true;

                    break;
            }
        }

        private void UndoButton_Click(object sender, RoutedEventArgs e)
        {
            if (_recorder == null)
            {
                return;
            }

            var temp = _recorder.AudioBuffer[_recordStartSlide.Index - 1];

            // disable undo since we allow only 1 step undo
            undoButton.IsEnabled = false;

            // revert back the last action
            _recorder.UndoLastRecord(_recordStartClick, _recordStartSlide);
            temp.RemoveAt(temp.Count - 1);

            // goto previous slide and click
            _slideShowWindow.View.GotoSlide(GetRelativeSlideIndex(_recordStartSlide.Index));
            _slideShowWindow.View.GotoClick(_recordStartClick);

            // activate the show window to allow user click on the slide show
            _slideShowWindow.Activate();
        }
    }
}
