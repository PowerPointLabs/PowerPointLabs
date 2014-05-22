using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using PowerPointLabs.AudioMisc;

namespace PowerPointLabs
{
    public partial class RecorderTaskPane : UserControl
    {
        // for hashing the speaker's script
        private MD5 _md5 = MD5.Create();

        // data structures to track embedded audio information
        
        // map the text MD5 to record id
        private Dictionary<string, int> _md5ScriptMapper;
        // map the scipt id to record id
        private Dictionary<int, int> _scriptRecrodMapper;
        // map a slide id to relative slide id
        private Dictionary<int, int> _relativeSlideIDmapper;
        // a collection of slides, each slide has a list of audio object
        private List<List<Audio>> _audioList;
        // a collection of slides, each slide has a list of script
        private List<List<string>> _scriptList;

        // slide monitor
        private PowerPointSlide _currentSlide;
        public int _currentSlideID;
        public bool _clearRecord = false;

        // Records save and display
        private readonly string _tempPath = Path.GetTempPath();
        private const string TempFolderName = "\\PowerPointLabs Temp\\";
        private const string SaveNameFormat = "Slide {0} Speech";
        private const string SpeechShapePrefix = "PowerPointLabs Speech";

        private enum RecorderStatus
        {
            Idle,
            Recording,
            Playing,
            Pause
        }
        private enum ScriptStatus
        {
            Generated,
            Recorded,
            None
        }

        # region Helper Functions
        /// <summary>
        /// This function will reset the UI to the default state.
        /// </summary>
        private void ResetRecorder()
        {
            soundTrackBar.Value = 0;
            timerLabel.Text = "00:00:00";
            statusLabel.Text = "Ready.";

            recButton.Text = "Record";
            stopButton.Text = "Stop";
            playButton.Text = "Play";

            _recButtonStatus = RecorderStatus.Idle;
            _playButtonStatus = RecorderStatus.Idle;
        }

        /// <summary>
        /// This function will dispose the timer and reset the timer count.
        /// </summary>
        private void ResetTimer()
        {
            _timerCnt = 0;
            timerLabel.Text = "00:00:00";
            if (_timer != null)
            {
                _timer.Dispose();
            }
        }

        /// <summary>
        /// This function will terminate the track bar event and set the track
        /// bar value to default position
        /// </summary>
        /// <param name="soundBarDefaultPos">Default position of the sound bar.</param>
        private void ResetTrackbar(int soundBarDefaultPos)
        {
            if (_trackbarThread != null && _trackbarThread.IsAlive)
            {
                _trackbarThread.Interrupt();
            }

            if (_stopwatch != null)
            {
                _stopwatch.Reset();
            }

            soundTrackBar.Value = soundBarDefaultPos;
        }

        /// <summary>
        /// This function will reset all unfinished session, including running
        /// timers and running sound.
        /// </summary>
        private void ResetSession()
        {
            // close unfinished sound session
            CloseAudio();

            // reset timer and trackbar
            ResetTimer();
            ResetTrackbar(0);
        }

        /// <summary>
        /// This function will convert a time in milli-second to HH:MM:SS:MMS
        /// </summary>
        /// <param name="millis">Time in millis.</param>
        /// <returns>A string in HH:MM:SS:MMS format.</returns>
        private string ConvertMillisToTime(long millis)
        {
            int ms, s, m, h;

            ms = (int)millis % 1000;
            millis /= 1000;

            s = (int)(millis % 60);
            millis /= 60;

            m = (int)(millis % 60);
            millis /= 60;

            h = (int)(millis % 60);
            millis /= 60;

            return System.String.Format("{0:D2}:{1:D2}:{2:D2}", h, m, s);
        }

        private string ConvertMillisToTime(int millis)
        {
            return ConvertMillisToTime((long)millis);
        }

        private string GetMD5(string s)
        {
            var hashcode = _md5.ComputeHash(System.Text.Encoding.UTF8.GetBytes(s));
            StringBuilder sb = new StringBuilder();

            foreach (byte x in hashcode)
            {
                sb.Append(x.ToString("X2"));
            }

            return sb.ToString();
        }

        private int GetRelativeSlideID(int curID)
        {
            if (!_relativeSlideIDmapper.ContainsKey(curID))
            {
                _relativeSlideIDmapper[curID] = _audioList.Count;
            }
            
            return _relativeSlideIDmapper[curID];
        }

        private void OpenNewAudio()
        {
            Native.mciSendString("open new type waveaudio alias sound", null, 0, IntPtr.Zero);
        }

        private void OpenAudio(string name)
        {
            Native.mciSendString("open \"" + name + "\" alias sound", null, 0, IntPtr.Zero);
        }

        private void CloseAudio()
        {
            Native.mciSendString("close sound", null, 0, IntPtr.Zero);
        }

        private int GetAudioLength()
        {
            Native.mciSendString("status sound length", mciRetInfo, MCI_RET_INFO_BUF_LEN, IntPtr.Zero);
            return Int32.Parse(mciRetInfo.ToString());
        }

        private int GetAudioLength(string name)
        {
            int length;

            OpenAudio(name);
            length = GetAudioLength();
            CloseAudio();

            return length;
        }

        private string GetAudioLengthString()
        {
            int length = GetAudioLength();
            return ConvertMillisToTime(length);
        }

        private string GetAudioLengthString(string name)
        {
            string length;

            OpenAudio(name);
            length = GetAudioLengthString();
            CloseAudio();

            return length;
        }

        private Audio.AudioType GetAudioType(string name)
        {
            if (name.Contains("Rec"))
            {
                return Audio.AudioType.Record;
            }

            return Audio.AudioType.Auto;
        }

        private bool LoadPlayback()
        {
            int selected = -1;

            // if no record is found, return false
            if (_curRecNumber == 0)
            {
                return false;
            }

            for (int i = 0; i < _curRecNumber; i++)
            {
                if (recDisplay.Items[i].Selected)
                {
                    selected = i;
                    break;
                }
            }

            if (selected == -1)
            {
                selected = _curRecNumber - 1;
                recDisplay.Items[selected].Selected = true;
            }

            _curPlayBack = _tempPath + "Rec" + selected.ToString() + ".wav";

            return true;
        }

        private void UpdateRecordList(string length)
        {
            // add the latest record to the list
            ListViewItem item = recDisplay.Items.Add(_curRecNumber.ToString());
            item.SubItems.Add("Rec" + _curRecNumber.ToString());
            item.SubItems.Add(length);
            item.SubItems.Add(DateTime.Now.ToString());

            // and select it by default
            recDisplay.Items[_curRecNumber - 1].Selected = true;
        }

        private void UpdateRecordList(int index, string name, string length)
        {
            // add the latest record to the list
            ListViewItem item = recDisplay.Items.Add(index.ToString());
            item.SubItems.Add(name);
            item.SubItems.Add(length);
            item.SubItems.Add(DateTime.Now.ToString());
        }

        private void UpdateScriptList(string name, ScriptStatus status)
        {
            ListViewItem item = scriptDisplay.Items.Add(status.ToString());
            item.SubItems.Add(name);
        }

        public void UpdateLists(int slideID)
        {
            int relativeID = GetRelativeSlideID(slideID);
            List<Audio> audio = _audioList[relativeID];
            List<string> scirpt = _scriptList[relativeID];

            // TODO:
            // Clear all + add all will be very slow, find some means to
            // do it faster

            // update the record list view
            ClearRecordList();
            recDisplay.BeginUpdate();
            for (int i = 0; i < audio.Count; i++)
            {
                UpdateRecordList(i, audio[i].SaveName, GetAudioLengthString(audio[i].SaveName));
            }
            recDisplay.EndUpdate();

            // update the script list view
            ClearScriptList();
            scriptDisplay.BeginUpdate();
            for (int i = 0; i < scirpt.Count; i++)
            {
                if (audio[i].Type == Audio.AudioType.Auto)
                {
                    UpdateScriptList(scirpt[i], ScriptStatus.Generated);
                }
                else
                {
                    UpdateScriptList(scirpt[i], ScriptStatus.Recorded);
                }
            }
            scriptDisplay.EndUpdate();
        }

        public void ClearRecordList()
        {
            recDisplay.BeginUpdate();
            recDisplay.Items.Clear();
            recDisplay.EndUpdate();
        }

        public void ClearScriptList()
        {
            scriptDisplay.BeginUpdate();
            scriptDisplay.Items.Clear();
            scriptDisplay.EndUpdate();
        }

        public void ClearLists()
        {
            ClearRecordList();
            ClearScriptList();
        }

        public void InitializeAudioAndScript(int slideID, string[] names, bool forceRefresh)
        {
            string[] audioSaveNames = null;
            string folderPath = Path.GetTempPath() + TempFolderName;
            int relativeSlideID = GetRelativeSlideID(slideID);
            _currentSlide = PowerPointPresentation.CurrentSlide;
            bool initialized = _audioList != null && _audioList.Count > relativeSlideID;

            // check if the selected slide has been initialized before
            if (initialized)
            {
                // TODO: 
                // if the slide has been initialized, check if the record has been updated

                // currently using forceRefresh to force an entire refresh
                if (!forceRefresh)
                {
                    return;
                }
            }

            // if the script of the selected slide has not been initialized yet,
            // we need to sniff the note pane to initialize the script list

            // TODO:
            // now we assume the first record -> first chunk of note, ect.

            // retrieve the tag notes
            var taggedNotes = new TaggedText(_currentSlide.NotesPageText.Trim());
            List<String> splitScript = taggedNotes.SplitByClicks();

            // if the slide has been initialized, update the list
            if (initialized)
            {
                _scriptList[relativeSlideID] = splitScript;
            }
            else
            // add the splitted notes into script list
            {
                _scriptList.Add(splitScript);
            }

            // map the md5 to script list index, this is used to do reorder
            for (int i = 0; i < splitScript.Count; i++)
            {
                string md5 = GetMD5(splitScript[i]);
                _md5ScriptMapper[md5] = i;
            }

            // if the audio of the selected slide has not been initialized yet,
            // we need to put all audio in the current slide into the list.
            if (!initialized)
            {
                _audioList.Add(new List<Audio>());
            }

            // if audio names have not been given, retrieve from files.
            if (names == null)
            {
                // retrieve all actual audio files in the slide
                String fileNameSearchPattern = String.Format(SaveNameFormat, _currentSlideID);
                var filePaths = Directory.EnumerateFiles(folderPath, "*.wav");
                audioSaveNames = filePaths.Where(path => path.Contains(fileNameSearchPattern)).ToArray();
            }
            else
            {
                audioSaveNames = names;
            }

            // construct audio object and put into audio collection
            for (int i = 0; i < audioSaveNames.Length; i++)
            {
                Audio audio = new Audio();
                string saveName = audioSaveNames[i];

                audio.SaveName = saveName;
                audio.Length = GetAudioLengthString(saveName);
                audio.LengthMillis = GetAudioLength(saveName);
                audio.MatchSciptID = i;
                audio.Type = GetAudioType(saveName);

                _audioList[relativeSlideID].Add(audio);
            }
        }

        public void InitializeAudioAndScript(List<string[]> names, bool forceRefresh)
        {
            // TODO:
            // if a slide has been initialized, check if some of the records have been updated
            // currently use forceRefresh to force an entire refresh
            var slides = PowerPointPresentation.Slides.ToList();

            for (int i = 0; i < slides.Count; i ++)
            {
                var slide = slides[i];

                InitializeAudioAndScript(slide.ID, names[i], forceRefresh);
            }
        }
        # endregion

        # region WinForm
        private const int MM_MCINOTIFY = 0x03B9;
        private const int MCI_NOTIFY_SUCCESS = 0x01;
        private const int MCI_NOTIFY_ABORTED = 0x04;
        private const int MCI_NOTIFY_FAILURE = 0x08;

        private const int MCI_RET_INFO_BUF_LEN = 128;

        private StringBuilder mciRetInfo;
        
        private string _curPlayBack = "";
        private int _resumeWaitingTime;
        private int _playbackLenMillis;
        private int _playbackTimeCnt;
        private int _timerCnt;

        private RecorderStatus _recButtonStatus;
        private RecorderStatus _playButtonStatus;

        private System.Threading.Timer _timer;
        private Thread _trackbarThread;

        private Stopwatch _stopwatch;
        
        private int _curRecNumber;

        // delgates to make thread safe control calls
        private delegate void SetLabelTextCallBack(Label label, string text);
        private delegate void SetTrackbarCallBack(TrackBar bar, int pos);
        private delegate void MCISendStringCallBack(string mciCommand,
                                                    StringBuilder mciRetInfo,
                                                    int infoLen,
                                                    IntPtr callBack);

        // delegates as notifiers
        public delegate void RecordStopNotify(string recName);
        public RecordStopNotify StopNotifier;

        // call when the pane becomes visible for the first time
        private void RecorderPane_Load(object sender, EventArgs e)
        {
            statusLabel.Text = "Ready.";
            statusLabel.Visible = true;
            _curRecNumber = 0;
            ResetRecorder();

            _currentSlide = PowerPointPresentation.CurrentSlide;
            if (_currentSlide != null)
            {
                UpdateLists(_currentSlide.ID);
            }
        }

        // call when the pane becomes visible from the second time onwards
        public void RecorderPaneReload()
        {
            statusLabel.Text = "Ready.";
            statusLabel.Visible = true;
            _curRecNumber = 0;
            ResetRecorder();

            _currentSlide = PowerPointPresentation.CurrentSlide;
            if (_currentSlide != null)
            {
                UpdateLists(_currentSlide.ID);
            }
        }

        // disable timer and thread when the pane is closed
        public void RecorderPaneClosing()
        {
            // before closing, clean up all unfinished sessions
            CloseAudio();

            if (_timer != null)
            {
                _timer.Dispose();
            }

            if (_trackbarThread != null && _trackbarThread.IsAlive)
            {
                _trackbarThread.Abort();
            }
        }

        # region Thread Safe Control Methods
        private void ThreadSafeUpdateLabelText(Label label, string time)
        {
            if (label.InvokeRequired)
            {
                SetLabelTextCallBack callback = new SetLabelTextCallBack(ThreadSafeUpdateLabelText);
                Invoke(callback, new object[] { label, time });
            }
            else
            {
                label.Text = time;
            }
        }

        private void ThreadSafeUpdateTrackbarValue(TrackBar bar, int value)
        {
            if (bar.InvokeRequired)
            {
                SetTrackbarCallBack callback = new SetTrackbarCallBack(ThreadSafeUpdateTrackbarValue);
                Invoke(callback, new object[] { bar, value });
            }
            else
            {
                int temp = (int) (value / (double) _playbackLenMillis * bar.Maximum);
                if (temp > bar.Maximum) temp = bar.Maximum;

                bar.Value = temp;
            }
        }

        private void ThreadSafeMCI(string mciCommand,
                                   StringBuilder mciRetInfo,
                                   int infoLen,
                                   IntPtr callBack)
        {
            if (this.InvokeRequired)
            {
                MCISendStringCallBack mciCallBack = new MCISendStringCallBack(ThreadSafeMCI);
                Invoke(mciCallBack, new object[]
                                        {
                                            mciCommand,
                                            mciRetInfo,
                                            infoLen,
                                            callBack
                                        });
            }
            else
            {
                Native.mciSendString(mciCommand,
                              mciRetInfo,
                              infoLen,
                              callBack);
            }
        }
        # endregion

        # region Timer and Trackbar Regualr Event Handlers
        private void TimerEvent(Object o)
        {
            ThreadSafeUpdateLabelText(timerLabel, ConvertMillisToTime(_timerCnt * 1000));
            _timerCnt++;
        }

        private void TrackbarEvent(Object o)
        {
            if (_stopwatch == null)
            {
                _stopwatch = Stopwatch.StartNew();
            }
            else
            {
                _stopwatch.Start();
            }

            try
            {
                while (true)
                {
                    if (_stopwatch.ElapsedMilliseconds % 5 == 0)
                    {
                        ThreadSafeUpdateTrackbarValue(soundTrackBar, (int)_stopwatch.ElapsedMilliseconds);
                    }
                }
            }
            catch (ThreadInterruptedException interrupt)
            {
            }
        }
        # endregion

        # region Button Event Handlers
        /// <summary>
        /// Handler handles the click event when the button is at idle state.
        /// Note: The routine will reset all other sessions.
        /// </summary>
        private void RecButtonIdleHandler()
        {
            // close unfinished session
            ResetSession();

            // UI settings
            ResetRecorder();
            statusLabel.Text = "Recording...";
            statusLabel.Visible = true;
            recButton.Text = "Pause";
            // disable control of playing
            playButton.Enabled = false;

            // change the status to recording status and change the button text
            // to pause
            _recButtonStatus = RecorderStatus.Recording;
            recButton.Text = "Pause";

            // start recording
            OpenNewAudio();
            Native.mciSendString("record sound", null, 0, IntPtr.Zero);

            // start the timer
            _timerCnt = 0;
            _timer = new System.Threading.Timer(TimerEvent, null, 0, 1000);
        }

        /// <summary>
        /// Handler handles the click event when the sound is recording. The
        /// recording will be paused and the timer will stop at the current
        /// length.
        /// </summary>
        private void RecButtonRecordingHandler()
        {
            // change the status to pause and change the button text to resume
            _recButtonStatus = RecorderStatus.Pause;
            recButton.Text = "Resume";

            // pause the sound and stop the timer
            _timer.Dispose();
            Native.mciSendString("pause sound", null, 0, IntPtr.Zero);

            // since the timer is counting in seconds, we need to know how many
            // millis to wait before next integral second.

            // retrieve current length
            int currentLen = GetAudioLength();
            _resumeWaitingTime = _timerCnt * 1000 - currentLen;

            if (_resumeWaitingTime < 0)
            {
                _resumeWaitingTime = 0;
            }
        }

        /// <summary>
        /// Handler handles click event when the sound recording is paused. The
        /// recording will resume and the timer will keep counting.
        /// </summary>
        private void RecButtonPauseHandler()
        {
            // change the status to recording and change the button text to
            // pause
            _recButtonStatus = RecorderStatus.Recording;
            recButton.Text = "Pause";

            // resume recording and restart the timer
            Native.mciSendString("resume sound", null, 0, IntPtr.Zero);
            _timer = new System.Threading.Timer(TimerEvent, null, _resumeWaitingTime, 1000);
        }

        /// <summary>
        /// Handler handles click event when sound is recording. It will save
        /// the sound to a user-specified path.
        /// </summary>
        private void StopButtonRecordingHandler()
        {
            // enable the control of play button
            playButton.Enabled = true;

            // change rec button status, rec button text, update status label
            // and stop timer
            _recButtonStatus = RecorderStatus.Idle;
            recButton.Text = "Record";
            statusLabel.Text = "Ready.";
            ResetTimer();

            // stop recording and get the length of the recording
            Native.mciSendString("stop sound", null, 0, IntPtr.Zero);
            // adjust the stop time difference between timer-stop and recording-stop
            timerLabel.Text = GetAudioLengthString();

            string saveName = _tempPath + "Rec" + _curRecNumber.ToString() + ".wav";
            _curRecNumber++;
            Native.mciSendString("save sound " + saveName, null, 0, IntPtr.Zero);
            CloseAudio();

            // update record list
            UpdateRecordList(timerLabel.Text);

            // notify outside
            StopNotifier(saveName);
        }

        /// <summary>
        /// Handler handles click event when the sound is playing back. It will
        /// stop the sound and reset all settings.
        /// </summary>
        private void StopButtonPlayingHandler()
        {
            // change play button status, update play button text, update
            // status label and reset all sessions
            Native.mciSendString("stop sound", null, 0, IntPtr.Zero);
            ResetSession();
            _playButtonStatus = RecorderStatus.Idle;
            playButton.Text = "Play";
            statusLabel.Text = "Ready.";
        }

        /// <summary>
        /// Handler handles click event when idle.
        /// </summary>
        private void PlayButtonIdleHandler()
        {
            // close unfinished session
            ResetSession();

            if (!LoadPlayback())
            {
                MessageBox.Show("No record to play back. Please record first.");
            }
            else
            {
                // UI settings
                ResetRecorder();
                statusLabel.Text = "Playing...";
                statusLabel.Visible = true;

                // change the button status and change the button text
                _playButtonStatus = RecorderStatus.Playing;
                playButton.Text = "Pause";

                // get play back length
                OpenAudio(_curPlayBack);
                _playbackLenMillis = GetAudioLength();

                // start the timer and track bar
                _playbackTimeCnt = 0;
                _timerCnt = 0;
                _timer = new System.Threading.Timer(TimerEvent, null, 0, 1000);
                _trackbarThread = new Thread(TrackbarEvent);
                _trackbarThread.Start();

                // start play back
                Native.mciSendString("play sound notify", null, 0, this.Handle);
            }
        }

        /// <summary>
        /// Handler handles click event when the sound is playing. It pauses
        /// the sound, timer and track bar.
        /// </summary>
        private void PlayButtonPlayingHandler()
        {
            // change the status to pause and change the text to resume
            _playButtonStatus = RecorderStatus.Pause;
            playButton.Text = "Resume";

            // pause the sound, timer and trackbar
            Native.mciSendString("pause sound", null, 0, IntPtr.Zero);
            _timer.Dispose();
            _stopwatch.Stop();
            _trackbarThread.Interrupt();

            // since the timer is counting in seconds, we need to know how many
            // millis to wait before next integral second.

            // retrieve current length
            Native.mciSendString("status sound position", mciRetInfo, MCI_RET_INFO_BUF_LEN, IntPtr.Zero);
            int currentLen = int.Parse(mciRetInfo.ToString());
            _resumeWaitingTime = _timerCnt * 1000 - currentLen;

            if (_resumeWaitingTime < 0)
            {
                _resumeWaitingTime = 0;
            }
        }

        /// <summary>
        /// Handler handles click event when the sound is paused. It resumes
        /// the sound, timer and track bar.
        /// </summary>
        private void PlayButtonPauseHandler()
        {
            // change the status to playing and change the button text to
            // pause
            _playButtonStatus = RecorderStatus.Playing;
            playButton.Text = "Pause";

            // resume recording, restart the timer and continue the track bar
            Native.mciSendString("resume sound", null, 0, IntPtr.Zero);
            _timer = new System.Threading.Timer(TimerEvent, null, _resumeWaitingTime, 1000);
            _trackbarThread = new Thread(TrackbarEvent);
            _trackbarThread.Start();
        }
        # endregion

        # region UI Control Events
        private void RecButtonClick(object sender, EventArgs e)
        {
            switch (_recButtonStatus)
            {
                case RecorderStatus.Idle:
                    RecButtonIdleHandler();
                    break;
                case RecorderStatus.Recording:
                    RecButtonRecordingHandler();
                    break;
                case RecorderStatus.Pause:
                    RecButtonPauseHandler();
                    break;
                default:
                    MessageBox.Show("Invalid Operation");
                    break;
            }
        }

        private void StopButtonClick(object sender, EventArgs e)
        {
            if (_recButtonStatus == RecorderStatus.Recording ||
                _recButtonStatus == RecorderStatus.Pause)
            {
                StopButtonRecordingHandler();
            } else
            if (_playButtonStatus == RecorderStatus.Playing ||
                _playButtonStatus == RecorderStatus.Pause)
            {
                StopButtonPlayingHandler();
            }
            else
            {
                MessageBox.Show("Invalid Operation");
            }
        }

        private void PlayButtonClick(object sender, EventArgs e)
        {
            switch (_playButtonStatus)
            {
                case RecorderStatus.Idle:
                    PlayButtonIdleHandler();
                    break;
                case RecorderStatus.Playing:
                    PlayButtonPlayingHandler();
                    break;
                case RecorderStatus.Pause:
                    PlayButtonPauseHandler();
                    break;
                default:
                    MessageBox.Show("Invalid Operation");
                    break;
            }
        }

        # endregion
        # endregion

        // do when the task pane first initialized
        public RecorderTaskPane()
        {
            mciRetInfo = new StringBuilder(MCI_RET_INFO_BUF_LEN);
            _audioList = new List<List<Audio>>();
            _scriptList = new List<List<string>>();
            
            _relativeSlideIDmapper = new Dictionary<int, int>();
            _scriptRecrodMapper = new Dictionary<int, int>();
            _md5ScriptMapper = new Dictionary<string, int>();
            
            InitializeComponent();
        }

        /// <summary>
        /// Overridden Win Form call back function, used to sniff call back
        /// messages triggered by MCI.
        /// </summary>
        /// <param name="m">A reference to the message sent by MCI.</param>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == MM_MCINOTIFY)
            {
                switch (m.WParam.ToInt32())
                {
                    case MCI_NOTIFY_SUCCESS:
                        // UI settings
                        statusLabel.Text = "Ready.";
                        playButton.Text = "Play";
                        _playButtonStatus = RecorderStatus.Idle;

                        // dispose timer and track bar timer while setting the
                        // track bar to full
                        ResetSession();
                        soundTrackBar.Value = soundTrackBar.Maximum;
                        break;
                    case MCI_NOTIFY_ABORTED:
                        ResetTrackbar(0);
                        break;
                    default:
                        MessageBox.Show("other error");
                        break;
                }
            }

            base.WndProc(ref m);
        }
    }
}
