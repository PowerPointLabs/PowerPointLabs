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
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PPExtraEventHelper;
using PowerPointLabs.Models;
using PowerPointLabs.AudioMisc;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{
    internal partial class RecorderTaskPane : UserControl
    {
        // for hashing the speaker's script
        private MD5 _md5 = MD5.Create();

        // data structures to track embedded audio information
        
        // map the text MD5 to record id
        private Dictionary<string, int> _md5ScriptMapper;
        // map the scipt id to record id
        private Dictionary<int, int> _scriptRecrodMapper;
        // this offset is used to map a slide id to relative slide id
        private const int slideIDOffset = 256;
        // a collection of slides, each slide has a list of audio object
        private List<List<Audio>> _audioList;
        // a collection of slides, each slide has a list of script
        private List<List<string>> _scriptList;

        // Records save and display
        private readonly string _tempPath = Path.GetTempPath();
        private const string TempFolderName = @"\PowerPointLabs Temp\";
        private readonly string tempFullPath = Path.GetTempPath() + TempFolderName;
        private const string SaveNameFormat = "Slide {0} Speech";
        private const string SpeechShapePrefix = "PowerPointLabs Speech";
        private const string SpeechShapeFormat = "PowerPointLabs Speech {0}";
        private const string ReopenSpeechFormat = "media{0}.wav";

        private enum RecorderStatus
        {
            Idle,
            Recording,
            Playing,
            Pause
        }
        private enum ScriptStatus
        {
            Default,
            Generated,
            Recorded,
            Untracked
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
            AudioHelper.CloseAudio();

            // reset timer and trackbar
            ResetTimer();
            ResetTrackbar(0);
        }

        private void SetAllRecorderButtonState(bool enable)
        {
            recButton.Enabled = enable;
            playButton.Enabled = enable;
            stopButton.Enabled = enable;
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

        private int GetRelativeSlideIndex(int curID)
        {
            return curID - slideIDOffset;
        }

        private Audio GetPlaybackFromList()
        {
            var slideID = GetRelativeSlideIndex(PowerPointPresentation.CurrentSlide.ID);
            int playbackIndex = -1;
            
            if (recDisplay.SelectedIndices.Count != 0)
            {
                playbackIndex = recDisplay.SelectedIndices[0];
            }
            
            if (playbackIndex == -1)
            {
                return null;
            }
            else
            {
                return _audioList[slideID][playbackIndex];
            }
        }

        private void UpdateRecordList(int index, string name, string length)
        {
            // change index to 1-base
            index ++;
            // add the latest record to the list
            if (index > recDisplay.Items.Count)
            {
                ListViewItem item = recDisplay.Items.Add(index.ToString());
                item.SubItems.Add(name);
                item.SubItems.Add(length);
                item.SubItems.Add(DateTime.Now.ToString());
            }
            else
            {
                // if name needs to be updated
                if (name != null)
                {
                    recDisplay.Items[index - 1].SubItems[1].Text = name;
                }

                // if length needs to be updated
                if (length != null)
                {
                    recDisplay.Items[index - 1].SubItems[2].Text = length;
                }

                recDisplay.Items[index - 1].SubItems[3].Text = DateTime.Now.ToString();
            }
        }

        private void UpdateScriptList(int index, string content, ScriptStatus status)
        {
            // change index to 1-base
            index++;

            if (index > scriptDisplay.Items.Count)
            {
                ListViewItem item = scriptDisplay.Items.Add(status.ToString());
                item.SubItems.Add(content);
            }
            else
            {
                if (status != ScriptStatus.Default)
                {
                    scriptDisplay.Items[index - 1].SubItems[0].Text = status.ToString();
                }

                if (content != null)
                {
                    scriptDisplay.Items[index - 1].SubItems[1].Text = content;
                }
            }
        }

        public void UpdateLists(int slideID)
        {
            int relativeID = GetRelativeSlideIndex(slideID);
            List<Audio> audio = _audioList[relativeID];
            List<string> scirpt = _scriptList[relativeID];

            // TODO:
            // Clear all + add all will be very slow, find some means to
            // do it faster

            // update the record list view
            ClearRecordDisplayList();
            recDisplay.BeginUpdate();
            for (int i = 0; i < audio.Count; i++)
            {
                UpdateRecordList(i, audio[i].Name, AudioHelper.GetAudioLengthString(audio[i].SaveName));
            }
            recDisplay.EndUpdate();

            // update the script list view
            ClearScriptDisplayList();
            scriptDisplay.BeginUpdate();
            for (int i = 0; i < scirpt.Count; i++)
            {
                if (audio[i].Type == Audio.AudioType.Auto)
                {
                    UpdateScriptList(i, scirpt[i], ScriptStatus.Generated);
                }
                else
                {
                    UpdateScriptList(i, scirpt[i], ScriptStatus.Recorded);
                }
            }
            scriptDisplay.EndUpdate();

            // since the pane was just renewed, no item is selected thus all
            // button should be disabled
            SetAllRecorderButtonState(false);
        }

        public void ClearRecordDisplayList()
        {
            recDisplay.BeginUpdate();
            recDisplay.Items.Clear();
            recDisplay.EndUpdate();
        }

        public void ClearScriptDisplayList()
        {
            scriptDisplay.BeginUpdate();
            scriptDisplay.Items.Clear();
            scriptDisplay.EndUpdate();
        }

        public void ClearDisplayLists()
        {
            ClearRecordDisplayList();
            ClearScriptDisplayList();
        }

        public void ClearRecordDataList()
        {
            foreach (var slide in _audioList)
            {
                slide.Clear();
            }
        }

        public void ClearRecordDataList(int id)
        {
            int relativeIndex = GetRelativeSlideIndex(id);
            _audioList[relativeIndex].Clear();
        }

        public void ClearScriptDataList()
        {
            foreach (var slide in _scriptList)
            {
                slide.Clear();
            }
        }

        public void ClearScriptDataList(int id)
        {
            int relativeIndex = GetRelativeSlideIndex(id);
            _scriptList[relativeIndex].Clear();
        }

        public void ClearDataLists()
        {
            ClearRecordDataList();
            ClearScriptDataList();
        }

        public void ClearDataLists(int id)
        {
            ClearRecordDataList(id);
            ClearScriptDataList(id);
        }

        public void SetupListsWhenOpen()
        {
            var slides = PowerPointPresentation.Slides.ToList();
            // track the total count of valid speech audio, this helps avoid
            // mixing up other audios with speech audios
            int validSpeechCnt = 0;
            
            foreach (var slide in slides)
            {
                // retrieve the tag notes
                var taggedNotes = new TaggedText(slide.NotesPageText.Trim());
                List<String> splitScript = taggedNotes.SplitByClicks();

                // add the splitted notes into script list
                _scriptList.Add(splitScript);

                // mapping the shapes with media files, and set up the audio list

                // append a new list of of audios to the current presentatoin audio list
                _audioList.Add(new List<Audio>());
                
                // get all audio shapes
                var shapes = slide.GetShapesWithMediaType(PpMediaType.ppMediaTypeSound);

                // iterate through all shapes, skip audios that are not generated speech
                for (int i = 0; i < shapes.Count; i++)
                {
                    var shape = shapes[i];

                    // if current audio is a speech, dump it into Audio object
                    if (shape.Name.Contains(SpeechShapePrefix))
                    {
                        var audio = new Audio();

                        // detect audio type
                        if (shape.MediaFormat.AudioSamplingRate == Audio.GeneratedSamplingRate)
                        {
                            audio.Type = Audio.AudioType.Auto;
                        }
                        else
                        if (shape.MediaFormat.AudioSamplingRate == Audio.RecordedSamplingRate)
                        {
                            audio.Type = Audio.AudioType.Record;
                        }
                        else
                        {
                            MessageBox.Show("Unrecognize Embedded Audio");
                        }

                        audio.SaveName = tempFullPath + String.Format(ReopenSpeechFormat, validSpeechCnt + 1);
                        audio.Name = String.Format(SpeechShapeFormat, validSpeechCnt);
                        audio.MatchSciptID = validSpeechCnt;
                        audio.Length = AudioHelper.GetAudioLengthString(audio.SaveName);
                        audio.LengthMillis = AudioHelper.GetAudioLength(audio.SaveName);

                        _audioList[slide.Index - 1].Add(audio);

                        validSpeechCnt++;
                    }
                }
            }
        }

        public void ShutdownReembed()
        {
            var slides = PowerPointPresentation.Slides.ToList();

            foreach (var slide in slides)
            {
                foreach (var audio in _audioList[slide.Index - 1])
                {
                    audio.EmbedOnSlide(slide);
                }
            }
        }

        public void InitializeAudioAndScript(PowerPointSlide slide, string[] names, bool forceRefresh)
        {
            string[] audioSaveNames = null;
            string folderPath = _tempPath + TempFolderName;
            
            int slideID = slide.ID;
            int relativeSlideID = GetRelativeSlideIndex(slideID);
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
            var taggedNotes = new TaggedText(slide.NotesPageText.Trim());
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
            // else clear the audio collection of current slide
            // TODO:
            // obviously we don't need to delete all items in the list, only
            // those modified items should be replaced.
            else
            {
                _audioList[relativeSlideID].Clear();
            }

            // if audio names have not been given, retrieve from files.
            if (names == null)
            {
                // retrieve all actual audio files in the slide
                String fileNameSearchPattern = String.Format(SaveNameFormat, slideID);
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
                string saveName = audioSaveNames[i];
                string name = String.Format(SpeechShapeFormat, i);
                var audio = new Audio(name, saveName, i);

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

                InitializeAudioAndScript(slide, names[i], forceRefresh);
            }
        }
        # endregion

        # region WinForm
        private int _resumeWaitingTime;
        private int _playbackLenMillis;
        private int _timerCnt;

        private RecorderStatus _recButtonStatus;
        private RecorderStatus _playButtonStatus;

        private System.Threading.Timer _timer;
        private Thread _trackbarThread;

        private Stopwatch _stopwatch;

        // delgates to make thread safe control calls
        private delegate void SetLabelTextCallBack(Label label, string text);
        private delegate void SetTrackbarCallBack(TrackBar bar, int pos);
        private delegate void MCISendStringCallBack(string mciCommand,
                                                    StringBuilder mciRetInfo,
                                                    int infoLen,
                                                    IntPtr callBack);

        // delegates as notifiers
        public delegate void RecordStopNotify(Audio audio);
        public RecordStopNotify ReplaceSoundShape;

        // call when the pane becomes visible for the first time
        private void RecorderPane_Load(object sender, EventArgs e)
        {
            statusLabel.Text = "Ready.";
            statusLabel.Visible = true;
            ResetRecorder();

            // disable all buttons when just enter the pane and nothing has
            // been selected
            SetAllRecorderButtonState(false);

            var currentSlide = PowerPointPresentation.CurrentSlide;
            if (currentSlide != null)
            {
                UpdateLists(currentSlide.ID);
            }
        }

        // call when the pane becomes visible from the second time onwards
        public void RecorderPaneReload()
        {
            statusLabel.Text = "Ready.";
            statusLabel.Visible = true;
            ResetRecorder();

            // disable record button when just enter the pane and nothing has
            // been selected
            SetAllRecorderButtonState(false);

            var currentSlide = PowerPointPresentation.CurrentSlide;
            if (currentSlide != null)
            {
                UpdateLists(currentSlide.ID);
            }
        }

        // disable timer and thread when the pane is closed
        public void RecorderPaneClosing()
        {
            // before closing, clean up all unfinished sessions
            AudioHelper.CloseAudio();

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
            ThreadSafeUpdateLabelText(timerLabel, AudioHelper.ConvertMillisToTime(_timerCnt * 1000));
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
            // enable stop button
            stopButton.Enabled = true;
            // disable control of both lists
            recDisplay.Enabled = false;
            scriptDisplay.Enabled = false;

            // change the status to recording status and change the button text
            // to pause
            _recButtonStatus = RecorderStatus.Recording;
            recButton.Text = "Pause";

            // start recording
            AudioHelper.OpenNewAudio();
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
            // make sure stop button is enabled
            stopButton.Enabled = true;

            // change the status to pause and change the button text to resume
            _recButtonStatus = RecorderStatus.Pause;
            statusLabel.Text = "Pause";
            recButton.Text = "Resume";

            // pause the sound and stop the timer
            _timer.Dispose();
            Native.mciSendString("pause sound", null, 0, IntPtr.Zero);

            // since the timer is counting in seconds, we need to know how many
            // millis to wait before next integral second.

            // retrieve current length
            int currentLen = AudioHelper.GetAudioLength();
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
            // make sure stop button is enabled
            stopButton.Enabled = true;

            // change the status to recording and change the button text to
            // pause
            _recButtonStatus = RecorderStatus.Recording;
            statusLabel.Text = "Recording...";
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

            var currentPlayback = GetPlaybackFromList();

            // stop recording and get the length of the recording
            Native.mciSendString("stop sound", null, 0, IntPtr.Zero);
            // adjust the stop time difference between timer-stop and recording-stop
            timerLabel.Text = AudioHelper.GetAudioLengthString();

            // ask if the user wants to do the replacement
            if (MessageBox.Show("Do you want to replace\n" + currentPlayback.SaveName + "\nwith current record?",
                                "Replacement", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                // user wants to do the replacement, save the file and replace the record
                string saveName = currentPlayback.SaveName.Replace(".wav", " rec.wav");
                Audio replaceRec = AudioHelper.DumpAudio(currentPlayback.Name, saveName, currentPlayback.MatchSciptID);

                // delete the old file
                File.Delete(currentPlayback.SaveName);

                // save curent sound
                Native.mciSendString("save sound \"" + saveName + "\"", null, 0, IntPtr.Zero);
                AudioHelper.CloseAudio();

                // update record list
                var relativeID = GetRelativeSlideIndex(PowerPointPresentation.CurrentSlide.ID);
                var replaceIndex = recDisplay.SelectedIndices[0];
                _audioList[relativeID][replaceIndex] = replaceRec;
                UpdateRecordList(replaceIndex, null, replaceRec.Length);

                // update the script list
                UpdateScriptList(replaceIndex, null, ScriptStatus.Recorded);

                // notify outside to embed the audio
                //ReplaceSoundShape(replaceRec);
                replaceRec.EmbedOnSlide(PowerPointPresentation.CurrentSlide);
            }

            // enable control of both lists
            recDisplay.Enabled = true;
            scriptDisplay.Enabled = true;
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

            // UI settings
            ResetSession();
            _playButtonStatus = RecorderStatus.Idle;
            playButton.Text = "Play";
            statusLabel.Text = "Ready.";
            // enable both lists
            recDisplay.Enabled = true;
            scriptDisplay.Enabled = true;
        }

        /// <summary>
        /// Handler handles click event when idle.
        /// </summary>
        private void PlayButtonIdleHandler()
        {
            // close unfinished session
            ResetSession();

            // UI settings
            ResetRecorder();
            statusLabel.Text = "Playing...";
            statusLabel.Visible = true;
            // enable stop button
            stopButton.Enabled = true;
            // disable control of both lists
            recDisplay.Enabled = false;
            scriptDisplay.Enabled = false;

            // change the button status and change the button text
            _playButtonStatus = RecorderStatus.Playing;
            playButton.Text = "Pause";

            // get play back length
            var playback = GetPlaybackFromList();

            // this shall not happen since the button is disabled when nothing
            // is selected
            if (playback == null)
            {
                MessageBox.Show("No record to play back. Please record first.");
            }
            else
            {
                _playbackLenMillis = playback.LengthMillis;

                // start the timer and track bar
                _timerCnt = 0;
                _timer = new System.Threading.Timer(TimerEvent, null, 0, 1000);
                _trackbarThread = new Thread(TrackbarEvent);
                _trackbarThread.Start();

                // start play back
                AudioHelper.OpenAudio(playback.SaveName);
                Native.mciSendString("play sound notify", null, 0, this.Handle);
            }
        }

        /// <summary>
        /// Handler handles click event when the sound is playing. It pauses
        /// the sound, timer and track bar.
        /// </summary>
        private void PlayButtonPlayingHandler()
        {
            // make sure stop button is enabled
            stopButton.Enabled = true;

            // change the status to pause and change the text to resume
            _playButtonStatus = RecorderStatus.Pause;
            statusLabel.Text = "Pause";
            playButton.Text = "Resume";

            // pause the sound, timer and trackbar
            Native.mciSendString("pause sound", null, 0, IntPtr.Zero);
            _timer.Dispose();
            _stopwatch.Stop();
            _trackbarThread.Interrupt();

            // since the timer is counting in seconds, we need to know how many
            // millis to wait before next integral second.

            // retrieve current length
            int currentLen = AudioHelper.GetAudioCurrentPosition();
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
            // make sure stop button is enabled
            stopButton.Enabled = true;

            // change the status to playing and change the button text to
            // pause
            _playButtonStatus = RecorderStatus.Playing;
            statusLabel.Text = "Playing...";
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

        private void RecDisplayItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            // if some record is selected, enable the record button
            if (e.IsSelected)
            {
                scriptDisplay.Items[e.ItemIndex].Selected = true;
                SetAllRecorderButtonState(true);
                stopButton.Enabled = false;
            }
            else
            {
                scriptDisplay.Items[e.ItemIndex].Selected = false;
                
                // disabling only happens when buttons are idle
                if (_playButtonStatus == RecorderStatus.Idle &&
                    _recButtonStatus == RecorderStatus.Idle)
                {
                    SetAllRecorderButtonState(false);
                }
            }
        }

        private void ScriptDisplayItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.IsSelected)
            {
                recDisplay.Items[e.ItemIndex].Selected = true;
                SetAllRecorderButtonState(true);
                stopButton.Enabled = false;
            }
            else
            {
                recDisplay.Items[e.ItemIndex].Selected = false;
                
                // disabling only happens when buttons are idle
                if (_playButtonStatus == RecorderStatus.Idle &&
                    _recButtonStatus == RecorderStatus.Idle)
                {
                    SetAllRecorderButtonState(false);
                }
            }
        }
        # endregion
        # endregion

        // do when the task pane first initialized
        public RecorderTaskPane()
        {
            _audioList = new List<List<Audio>>();
            _scriptList = new List<List<string>>();
            
            _scriptRecrodMapper = new Dictionary<int, int>();
            _md5ScriptMapper = new Dictionary<string, int>();
            
            InitializeComponent();

            // don't allow user to touch trackbar, thus disabled
            soundTrackBar.Enabled = false;
        }

        /// <summary>
        /// Overridden Win Form call back function, used to sniff call back
        /// messages triggered by MCI.
        /// </summary>
        /// <param name="m">A reference to the message sent by MCI.</param>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == AudioHelper.MM_MCINOTIFY)
            {
                switch (m.WParam.ToInt32())
                {
                    case AudioHelper.MCI_NOTIFY_SUCCESS:
                        // UI settings
                        statusLabel.Text = "Ready.";
                        playButton.Text = "Play";
                        _playButtonStatus = RecorderStatus.Idle;
                        // disable stop button
                        stopButton.Enabled = false;
                        // enable both lists
                        recDisplay.Enabled = true;
                        scriptDisplay.Enabled = true;

                        // dispose timer and track bar timer while setting the
                        // track bar to full
                        ResetSession();
                        soundTrackBar.Value = soundTrackBar.Maximum;
                        break;
                    case AudioHelper.MCI_NOTIFY_ABORTED:
                        ResetTrackbar(0);
                        break;
                    default:
                        MessageBox.Show("Fatal error");
                        break;
                }
            }

            base.WndProc(ref m);
        }
    }
}
