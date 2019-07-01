using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Tools;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.CaptionsLab;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Views;
using PowerPointLabs.FunctionalTestInterface.Impl;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
using PowerPointLabs.Models;
using PowerPointLabs.PositionsLab;
using PowerPointLabs.ResizeLab;
using PowerPointLabs.SaveLab;
using PowerPointLabs.ShapesLab;
using PowerPointLabs.SyncLab.Views;
using PowerPointLabs.TextCollection;
using PowerPointLabs.TimerLab;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using PPExtraEventHelper;

using MessageBox = System.Windows.Forms.MessageBox;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    public partial class ThisAddIn
    {
#pragma warning disable 0618
        public readonly int MaxCustomTaskPanes = 2;
        public readonly string OfficeVersion2013 = "15.0";
        public readonly string OfficeVersion2010 = "14.0";

        public static string AppDataFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PowerPointLabs");

        public Ribbon1 Ribbon;

        internal ShapesLabConfigSaveFile ShapesLabConfig;

        internal PowerPointShapeGalleryPresentation ShapePresentation;

        private delegate void SyncElearningItemsDelegate();

        private const string AppLogName = "PowerPointLabs.log";
        private const string SlideXmlSearchPattern = @"slide(\d+)\.xml";
        private const string TempFolderNamePrefix = @"\PowerPointLabs Temp\";
        private const string ShapeGalleryPptxName = "ShapeGallery";
        private const string SyncLabPptxName = "Sync Lab - Do not edit";
        private const string TempZipName = "tempZip.zip";

        private string _deactivatedPresFullName;
        private string tempFolderName;

        private bool _pptLabsShouldTerminate;

        private bool isResizePaneVisible;

        private readonly Dictionary<PowerPoint.DocumentWindow,
            List<CustomTaskPane>> _documentPaneMapper = new Dictionary<PowerPoint.DocumentWindow, List<CustomTaskPane>>();

        private readonly Dictionary<PowerPoint.DocumentWindow,
            string> _documentHashcodeMapper = new Dictionary<PowerPoint.DocumentWindow, string>();

        /// <summary>
        /// The channel for .NET Remoting calls.
        /// </summary>
        private IChannel _ftChannel;

        #region API

        public bool IsApplicationVersion2010()
        {
            return Application.Version == OfficeVersion2010;
        }

        public bool IsApplicationVersion2013()
        {
            return Application.Version == OfficeVersion2013;
        }

        public Control GetActiveControl(Type type)
        {
            CustomTaskPane taskPane = GetActivePane(type);

            return taskPane == null ? null : taskPane.Control;
        }

        public string GetTempFolderName()
        {
            return tempFolderName;
        }

        public CustomTaskPane GetActivePane(Type type)
        {
            PowerPoint.DocumentWindow activeWindow;
            try
            {
                activeWindow = Application.ActiveWindow;
            }
            catch (COMException)
            {
                return null;
            }
            return GetPaneFromWindow(type, activeWindow);
        }

        public Control GetControlFromWindow(Type type, PowerPoint.DocumentWindow window)
        {
            CustomTaskPane taskPane = GetPaneFromWindow(typeof(CustomShapePane), window);

            return taskPane == null ? null : taskPane.Control;
        }

        public CustomTaskPane GetPaneFromWindow(Type type, PowerPoint.DocumentWindow window)
        {
            if (!_documentPaneMapper.ContainsKey(window))
            {
                return null;
            }

            List<CustomTaskPane> panes = _documentPaneMapper[window];

            foreach (CustomTaskPane pane in panes)
            {
                try
                {
                    UserControl control = pane.Control;

                    if (control.GetType() == type)
                    {
                        return pane;
                    }
                }
                catch (Exception)
                {
                    return null;
                }
            }

            return null;
        }

        public string GetActiveWindowTempName()
        {
            return _documentHashcodeMapper[Application.ActiveWindow];
        }

        public void InitializeShapeGallery()
        {
            // achieves singleton ShapePresentation
            if (ShapePresentation?.Opened ?? false)
            {
                return;
            }

            string shapeRootFolderPath = ShapesLabSettings.SaveFolderPath;

            ShapePresentation =
                new PowerPointShapeGalleryPresentation(shapeRootFolderPath, ShapeGalleryPptxName);

            if (!ShapePresentation.Open(withWindow: false, focus: false) &&
                !ShapePresentation.Opened)
            {
                // if the presentation gets some error during opening, and the error could not
                // be resolved by consistency check, prompt the user about the error
                MessageBox.Show(CommonText.ErrorShapeGalleryInit);
                return;
            }

            if (ShapePresentation.HasCategory(ShapesLabConfig.DefaultCategory))
            {
                ShapePresentation.DefaultCategory = ShapesLabConfig.DefaultCategory;

                return;
            }

            // if we do not have the default category, create and add it to ShapeGallery
            ShapePresentation.AddCategory(ShapesLabConfig.DefaultCategory);
            ShapePresentation.Save();
        }

        public void InitializeShapesLabConfig()
        {
            // if ShapesLabConfig has already been intialized, do nothing
            if (ShapesLabConfig != null)
            {
                return;
            }

            ShapesLabConfig = new ShapesLabConfigSaveFile(AppDataFolder);

            // create a directory under specified location if the location does not exist
            if (!Directory.Exists(ShapesLabSettings.SaveFolderPath))
            {
                Directory.CreateDirectory(ShapesLabSettings.SaveFolderPath);
            }
        }

        public void PrepareMediaFiles(PowerPoint.Presentation pres, string tempPath)
        {
            string presFullName = pres.FullName;
            string presName = pres.Name;

            // in case of embedded slides, we need to regulate the file name and full name
            RegulatePresentationName(pres, tempPath, ref presName, ref presFullName);

            try
            {
                if (IsEmptyFile(presFullName))
                {
                    return;
                }

                string zipFullPath = tempPath + TempZipName;

                // before we do everything, check if there's an undelete old zip file
                // due to some error
                try
                {
                    FileDir.DeleteFile(zipFullPath);
                    FileDir.CopyFile(presFullName, zipFullPath);
                }
                catch (Exception e)
                {
                    ErrorDialogBox.ShowDialog(CommonText.ErrorAccessTempFolder, string.Empty, e);
                }

                ExtractMediaFiles(zipFullPath, tempPath);
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(CommonText.ErrorPrepareMedia, "Files cannot be linked.", e);
            }
        }

        public string PrepareTempFolder(PowerPoint.Presentation pres)
        {
            string presName = pres.Name;
            string presFullName = pres.FullName;

            // here presFullName makes no use, just to fit in the signature
            RegulatePresentationName(pres, null, ref presName, ref presFullName);

            string tempPath = GetPresentationTempFolder(presName);

            // if temp folder doesn't exist, create
            try
            {
                if (Directory.Exists(tempPath))
                {
                    Directory.Delete(tempPath, true);
                }
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(CommonText.ErrorCreateTempFolder, string.Empty, e);
            }
            finally
            {
                Directory.CreateDirectory(tempPath);
            }

            return tempPath;
        }

        public void RegisterResizePane(PowerPoint.Presentation presentation)
        {
            if (GetActivePane(typeof(ResizeLabPane)) != null)
            {
                return;
            }

            PowerPoint.DocumentWindow activeWindow = presentation.Application.ActiveWindow;

            RegisterTaskPane(new ResizeLabPane(), ResizeLabText.TaskPaneTitle, activeWindow,
                ResizeTaskPaneVisibleValueChangedEventHandler, null);
        }

        public void RegisterRecorderPane(PowerPoint.DocumentWindow activeWindow, string tempFullPath)
        {
            if (GetActivePane(typeof(RecorderTaskPane)) != null)
            {
                return;
            }

            RegisterTaskPane(new RecorderTaskPane(tempFullPath), NarrationsLabText.RecManagementPanelTitle, activeWindow,
                TaskPaneVisibleValueChangedEventHandler, null);
        }

        public void SyncShapeAdd(string shapeName, string shapeFullName, string category)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow)
                {
                    continue;
                }

                CustomShapePane shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl?.CurrentCategory == category)
                {
                    shapePaneControl.AddCustomShape(shapeName, shapeFullName, false);
                }
            }
        }

        public void SyncShapeRemove(string shapeName, string category)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow)
                {
                    continue;
                }

                CustomShapePane shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl?.CurrentCategory == category)
                {
                    shapePaneControl.RemoveCustomShape(shapeName);
                }
            }
        }

        public void SyncShapeRename(string shapeOldName, string shapeNewName, string category)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow)
                {
                    continue;
                }

                CustomShapePane shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl?.CurrentCategory == category)
                {
                    shapePaneControl.RenameCustomShape(shapeOldName, shapeNewName);
                }
            }
        }

        public bool VerifyOnLocal(PowerPoint.Presentation pres)
        {
            Regex invalidPathRegex = new Regex("^[hH]ttps?:");

            return !invalidPathRegex.IsMatch(pres.Path);
        }

        public bool VerifyVersion(PowerPoint.Presentation pres)
        {
            return !pres.Name.EndsWith(".ppt");
        }

        #endregion

        #region Helper Functions

        public CustomTaskPane RegisterTaskPane(UserControl control, string title, PowerPoint.DocumentWindow wnd,
    EventHandler visibleChangeEventHandler = null,
    EventHandler dockPositionChangeEventHandler = null)
        {
            LoadingDialogBox loadingDialog = new LoadingDialogBox();
            loadingDialog.Show();

            // note down the control's width
            int width = control.Width;

            // register the user control to the CustomTaskPanes collection and set it as
            // current active task pane;
            CustomTaskPane taskPane = CustomTaskPanes.Add(control, title, wnd);

            // task pane UI setup
            taskPane.Visible = false;
            taskPane.Width = width + 20;

            // map the current window with the task pane
            if (!_documentPaneMapper.ContainsKey(wnd))
            {
                _documentPaneMapper[wnd] = new List<CustomTaskPane>();
            }

            _documentPaneMapper[wnd].Add(taskPane);

            if (_documentPaneMapper[wnd].Count > MaxCustomTaskPanes)
            {
                // remove the oldest task pane
                Type oldestPaneType = _documentPaneMapper[wnd].First().Control.GetType();
                RemoveTaskPane(wnd, oldestPaneType);
                Trace.TraceInformation(
                    $"Removed pane {oldestPaneType.ToString()} over limit: {MaxCustomTaskPanes.ToString()}");
            }

            Trace.TraceInformation(
                "After Pane Width Change: " +
                string.Format("Pane Width = {0}, Pane Height = {1}, Control Width = {2}, Control Height {3}",
                    taskPane.Width, taskPane.Height, control.Width, control.Height));

            // event handlers register
            if (visibleChangeEventHandler != null)
            {
                taskPane.VisibleChanged += visibleChangeEventHandler;
            }

            if (dockPositionChangeEventHandler != null)
            {
                taskPane.DockPositionChanged += dockPositionChangeEventHandler;
            }

            loadingDialog.Close();
            return taskPane;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon = new Ribbon1();
            return Ribbon;
        }

        private void SetupLogger()
        {
            // Check if folder exists and if not, create it
            if (!Directory.Exists(AppDataFolder))
            {
                Directory.CreateDirectory(AppDataFolder);
            }

            string fileName = DateTime.Now.ToString("yyyy-MM-dd") + AppLogName;
            string logPath = Path.Combine(AppDataFolder, fileName);

            Trace.AutoFlush = true;
            Trace.Listeners.Add(new TextWriterTraceListener(logPath));
        }

        private void ShutDownRecorderPane()
        {
            RecorderTaskPane recorder = GetActiveControl(typeof(RecorderTaskPane)) as RecorderTaskPane;

            if (recorder?.HasEvent() ?? false)
            {
                recorder.ForceStopEvent();
            }
        }

        private void ShutDownPictureSlidesLab()
        {
            PictureSlidesLab.Views.PictureSlidesLabWindow pictureSlidesLabWindow = Ribbon.PictureSlidesLabWindow;
            if (pictureSlidesLabWindow?.IsOpen ?? false)
            {
                pictureSlidesLabWindow.Close();
            }
        }

        private void RemoveTaskPanes(PowerPoint.DocumentWindow activeWindow)
        {
            if (!_documentPaneMapper.ContainsKey(activeWindow))
            {
                return;
            }

            List<CustomTaskPane> activePanes = _documentPaneMapper[activeWindow];
            foreach (CustomTaskPane pane in activePanes)
            {
                CustomTaskPanes.Remove(pane);
            }

            _documentPaneMapper.Remove(activeWindow);
        }

        private void RemoveTaskPane(PowerPoint.DocumentWindow window, Type paneType)
        {
            if (!_documentPaneMapper.ContainsKey(window))
            {
                return;
            }

            List<CustomTaskPane> activePanes = _documentPaneMapper[window];
            for (int i = activePanes.Count - 1; i >= 0; i--)
            {
                CustomTaskPane pane = activePanes[i];
                if (pane.Control.GetType() != paneType)
                {
                    continue;
                }

                CustomTaskPanes.Remove(pane);
                activePanes.RemoveAt(i);
            }
        }

        private void RegulatePresentationName(PowerPoint.Presentation pres, string tempPath, ref string presName,
            ref string presFullName)
        {
            // this function is used to handle "embed on other application" issue. In this case,
            // all of presentation name, path and full name do not match the usual rule: name is
            // "Untitled", path is empty string and full name is "slide in XX application". We need
            // to regulate these fields properly.

            if (!presName.Contains(".pptx"))
            {
                presName += ".pptx";
            }

            if (tempPath != null)
            {
                // every time when recorder pane is open,
                // save this presentation's copy, which will be used
                // to load audio files later
                pres.SaveCopyAs(tempPath + presName);
                presFullName = tempPath + presName;
            }
        }

        private void TaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            CustomTaskPane recorderPane = GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            RecorderTaskPane recorder = recorderPane.Control as RecorderTaskPane;

            // trigger close form event when closing hide the pane
            if (!recorder?.Visible ?? false)
            {
                recorder.RecorderPaneClosing();
                // remove recorder pane and force it to reload when next time open
                RemoveTaskPane(Application.ActiveWindow, typeof(RecorderTaskPane));
            }
        }

        private void ResizeTaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            CustomTaskPane resizePane = GetActivePane(typeof(ResizeLabPane));

            if (resizePane == null)
            {
                return;
            }

            isResizePaneVisible = resizePane.Visible;
        }

        private bool SlidesInRangeHaveCaptions(PowerPoint.SlideRange sldRange)
        {
            foreach (PowerPoint.Slide slide in sldRange)
            {
                PowerPointSlide pptSlide = PowerPointSlide.FromSlideFactory(slide);
                if (pptSlide.HasCaptions())
                {
                    return true;
                }
            }
            return false;
        }

        private bool SlidesInRangeHaveAudio(PowerPoint.SlideRange sldRange)
        {
            foreach (PowerPoint.Slide slide in sldRange)
            {
                PowerPointSlide pptSlide = PowerPointSlide.FromSlideFactory(slide);
                if (pptSlide.HasAudio())
                {
                    return true;
                }
            }
            return false;
        }

        private void SlideShowBeginHandler(PowerPoint.SlideShowWindow wn)
        {
            _isInSlideShow = true;
            AgendaLab.AgendaLabMain.SlideShowBeginHandler();
        }

        private void SlideShowEndHandler(PowerPoint.Presentation presentation)
        {
            _isInSlideShow = false;

            RecorderTaskPane recorder = GetActiveControl(typeof(RecorderTaskPane)) as RecorderTaskPane;

            if (recorder == null)
            {
                AgendaLab.AgendaLabMain.SlideShowEndHandler();
                return;
            }

            // force recording session ends
            if (recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }

            // enable slide show button
            recorder.EnableSlideShow();

            // when leave the show, dispose the in-show control if we have one
            recorder.DisposeInSlideControlBox();

            // if audio buffer is not empty, render the effects
            if (recorder.AudioBuffer.Count != 0)
            {
                List<PowerPointSlide> slides = PowerPointPresentation.Current.Slides.ToList();

                for (int i = 0; i < recorder.AudioBuffer.Count; i++)
                {
                    if (recorder.AudioBuffer[i].Count == 0)
                    {
                        continue;
                    }

                    foreach (Tuple<AudioMisc.Audio, int> audio in recorder.AudioBuffer[i])
                    {
                        audio.Item1.EmbedOnSlide(slides[i], audio.Item2);

                        if (ELearningLab.Service.ComputerVoiceRuntimeService.IsRemoveAudioEnabled)
                        {
                            continue;
                        }

                        ELearningLab.Service.ComputerVoiceRuntimeService.IsRemoveAudioEnabled = true;
                        Ribbon.RefreshRibbonControl("RemoveNarrationsButton");
                    }
                }
            }

            // clear the buffer after embed
            recorder.AudioBuffer.Clear();

            // change back the slide range settings
            Application.ActivePresentation.SlideShowSettings.RangeType = PowerPoint.PpSlideShowRangeType.ppShowAll;

            AgendaLab.AgendaLabMain.SlideShowEndHandler();
        }

        private bool IsEmptyFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }

            FileInfo fileInfo = new FileInfo(filePath);

            return fileInfo.Length == 0;
        }

        private void UpdateRecorderPane(int count, int id)
        {
            CustomTaskPane recorderPane = GetActivePane(typeof(RecorderTaskPane));

            // if there's no active pane associated with the current window, return
            if (recorderPane == null)
            {
                return;
            }

            RecorderTaskPane recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null)
            {
                return;
            }

            // if the user has selected none or more than 1 slides, recorder pane should show nothing
            if (count != 1)
            {
                if (recorderPane.Visible)
                {
                    recorder.ClearDisplayLists();
                }
            }
            else
            {
                // initailize the current slide
                recorder.InitializeAudioAndScript(PowerPointCurrentPresentationInfo.CurrentSlide, null, false);

                // if the pane is shown, refresh the pane immediately
                if (recorderPane.Visible)
                {
                    recorder.UpdateLists(id);
                }
            }
        }

        private void UpdateTimerPane(bool isVisible)
        {
            CustomTaskPane timerPane = GetActivePane(typeof(TimerPane));
            if (timerPane != null)
            {
                timerPane.Visible = isVisible;
            }
        }

        private void UpdateELearningPane(int selectedSlidesCount)
        {
            CustomTaskPane elearningLabPane = GetActivePane(typeof(ELearningLabTaskpane));
            if (elearningLabPane == null)
            {
                return;
            }
            ELearningLabTaskpane taskpane = elearningLabPane.Control as ELearningLabTaskpane;
            taskpane.ELearningLabMainPanel.SyncElearningLabOnSlideSelectionChanged();
            if (elearningLabPane.Visible == true)
            {
                taskpane.ELearningLabMainPanel.ReloadELearningLabOnSlideSelectionChanged();
            }
        }

        private string GetPresentationTempFolder(string presName)
        {
            string tempName = presName.GetHashCode().ToString(CultureInfo.InvariantCulture);
            string tempPath = Path.GetTempPath() + TempFolderNamePrefix + tempName + @"\";

            return tempPath;
        }

        private void CleanUp(PowerPoint.DocumentWindow associatedWindow)
        {
            if (_documentHashcodeMapper.ContainsKey(associatedWindow))
            {
                _documentHashcodeMapper.Remove(associatedWindow);
            }

            // if there exists some task panes, remove them
            RemoveTaskPanes(associatedWindow);
        }

        private void ExtractMediaFiles(string zipFullPath, string tempPath)
        {
            try
            {
                ZipStorer zip = ZipStorer.Open(zipFullPath, FileAccess.Read);
                List<ZipStorer.ZipFileEntry> dir = zip.ReadCentralDir();

                Regex regex = new Regex(SlideXmlSearchPattern);

                foreach (ZipStorer.ZipFileEntry entry in dir)
                {
                    string name = Path.GetFileName(entry.FilenameInZip);

                    if (name == null)
                    {
                        continue;
                    }

                    if (name.Contains(".wav") ||
                        regex.IsMatch(name))
                    {
                        zip.ExtractFile(entry, tempPath + name);
                    }
                }

                zip.Close();

                FileDir.DeleteFile(zipFullPath);
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog(CommonText.ErrorExtract, "Archived files cannot be retrieved.", e);
            }
        }

        private void BreakRecorderEvents()
        {
            RecorderTaskPane recorder = GetActiveControl(typeof(RecorderTaskPane)) as RecorderTaskPane;

            if (recorder?.HasEvent() ?? false)
            {
                recorder.ForceStopEvent();
            }
        }

        private void ShutDownSyncLab()
        {
            // If sync lab open, then close it.
            PowerPoint.Presentation syncLabPpt = GetOpenedSyncLabPresentation();
            if (syncLabPpt != null)
            {
                syncLabPpt.Close();
                Trace.TraceInformation("SyncLab terminated.");
            }
        }

        private PowerPoint.Presentation GetOpenedSyncLabPresentation()
        {
            foreach (PowerPoint.Presentation presentation in Application.Presentations)
            {
                if (presentation.Name.Contains(SyncLabPptxName))
                {
                    return presentation;
                }
            }
            return null;
        }

        private void ShutDownShapesLab()
        {
            if (ShapePresentation?.Opened ?? false)
            {
                if (string.IsNullOrEmpty(ShapesLabConfig.DefaultCategory))
                {
                    ShapesLabConfig.DefaultCategory = ShapePresentation.Categories[0];
                }

                ShapePresentation.Close();
                Trace.TraceInformation("ShapesLab terminated.");
            }
        }
        # endregion

        # region Powerpoint Application Event Handlers

        private void ThisAddInStartup(object sender, EventArgs e)
        {
            SetupLogger();
            Logger.Log("PowerPointLabs Started");

            CultureUtil.SetDefaultCulture(CultureInfo.GetCultureInfo("en-US"));

            new Updater().TryUpdate();
            SetupFunctionalTestChannels();

            PPMouse.Init(Application);
            PPKeyboard.Init(Application);
            PPCopy.Init(Application);
            UIThreadExecutor.Init();
            SetupDoubleClickHandler();
            SetupTabActivateHandler();
            SetupAfterCopyPasteHandler();

            SaveLabSettings.InitialiseLocalStorage();
            PPLClipboard.Init(new IntPtr(Application.HWND));

            // According to MSDN, when more than 1 event are triggered, callback's invoking sequence
            // follows the defining order. I.e. the earlier you defined, the earlier it will be
            // executed.

            // Here, we want the priority to be: Application action > Window action > Slide action

            // Priority High: Application Actions
            ((PowerPoint.EApplication_Event)Application).NewPresentation += ThisAddInNewPresentation;
            Application.AfterNewPresentation += ThisAddInAfterNewPresentation;
            Application.PresentationOpen += ThisAddInPresentationOpen;
            Application.PresentationClose += ThisAddInPresentationClose;

            // Priority Mid: Window Actions
            Application.WindowActivate += ThisAddInApplicationOnWindowActivate;
            Application.WindowDeactivate += ThisAddInApplicationOnWindowDeactivate;
            Application.WindowSelectionChange += ThisAddInSelectionChanged;
            Application.SlideShowBegin += SlideShowBeginHandler;
            Application.SlideShowEnd += SlideShowEndHandler;

            // Priority Low: Slide Actions
            Application.SlideSelectionChanged += ThisAddInSlideSelectionChanged;
        }

        private void ThisAddInApplicationOnWindowDeactivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            Trace.TraceInformation(pres.Name + " (Presentation) and " + wn.Caption + " (Window) deactivated.");
            _deactivatedPresFullName = pres.FullName;
        }

        private void ThisAddInApplicationOnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            if (pres == null)
            {
                return;
            }
            Trace.TraceInformation(pres.Name + " (Presentation) and " + wn.Caption + " (Window) activated.");

            CustomShapePane customShape = GetActiveControl(typeof(CustomShapePane)) as CustomShapePane;

            // make sure ShapeGallery's default category is consistent with current presentation
            if (customShape != null)
            {
                string currentCategory = customShape.CurrentCategory;
                ShapePresentation.DefaultCategory = currentCategory;
            }

            // If a window was activated in any way, PptLabs should not terminate.
            _pptLabsShouldTerminate = false;
        }

        private void ThisAddInSlideSelectionChanged(PowerPoint.SlideRange sldRange)
        {
            // TODO: doing range sweep to check these var may affect performance, consider initializing these
            // TODO: variables only at program starts
            NotesToCaptions.IsRemoveCaptionsEnabled = SlidesInRangeHaveCaptions(sldRange);
            ComputerVoiceRuntimeService.IsRemoveAudioEnabled = SlidesInRangeHaveAudio(sldRange);

            // update recorder pane
            if (sldRange.Count > 0)
            {
                int slideID;
                try
                {
                    slideID = sldRange[1].SlideID;
                }
                catch (COMException)
                {
                    return;
                }

                UpdateRecorderPane(sldRange.Count, slideID);
                TimerLab.TimerLab.IsTimerEnabled = true;
                PictureSlidesLab.PictureSlidesLab.IsPictureSlidesEnabled = true;
            }
            else
            {
                UpdateRecorderPane(sldRange.Count, -1);
                TimerLab.TimerLab.IsTimerEnabled = false;
                UpdateTimerPane(false);
                PictureSlidesLab.PictureSlidesLab.IsPictureSlidesEnabled = false;
                ShutDownPictureSlidesLab();
            }

            UpdateELearningPane(sldRange.Count);
            // in case the recorder is on event
            BreakRecorderEvents();

            // ribbon function init
            HighlightLab.HighlightBulletsText.IsHighlightPointsEnabled = true;
            HighlightLab.HighlightBulletsBackground.IsHighlightBackgroundEnabled = true;

            if (sldRange.Count != 1)
            {
                HighlightLab.HighlightBulletsText.IsHighlightPointsEnabled = false;
                HighlightLab.HighlightBulletsBackground.IsHighlightBackgroundEnabled = false;
            }
            else
            {
                PowerPoint.Slide tmp = sldRange[1];
                PowerPoint.Presentation presentation = PowerPointPresentation.Current.Presentation;
                int slideIndex = tmp.SlideIndex;
                PowerPoint.Slide next = tmp;
                PowerPoint.Slide prev = tmp;

                if (slideIndex < presentation.Slides.Count)
                {
                    next = presentation.Slides[slideIndex + 1];
                }

                if (slideIndex > 1)
                {
                    prev = presentation.Slides[slideIndex - 1];
                }
            }
            Ribbon.RefreshRibbonControl("PictureSlidesLabButton");
            Ribbon.RefreshRibbonControl("TimerLabButton");
            Ribbon.RefreshRibbonControl("HighlightPointsButton");
            Ribbon.RefreshRibbonControl("HighlightBackgroundButton");
            Ribbon.RefreshRibbonControl("RemoveCaptionsButton");
            Ribbon.RefreshRibbonControl("RemoveNarrationsButton");
            Ribbon.RefreshRibbonControl("ELearningTaskPaneButton");
        }

        // To handle AccessViolationException
        [HandleProcessCorruptedStateExceptions]
        private void ThisAddInSelectionChanged(PowerPoint.Selection sel)
        {
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape sh = null;
                try
                {
                    sh = sel.ShapeRange[1];
                }
                catch (AccessViolationException e)
                {
                    Logger.LogException(e, "ThisAddInSelectionChanged");
                    Logger.Log("We do not have access to the ShapeRange now!");
                    return;
                }

                if (isResizePaneVisible)
                {
                    sel.ShapeRange.LockAspectRatio = ResizeLabPaneWPF.IsAspectRatioLocked
                        ? Office.MsoTriState.msoTrue
                        : Office.MsoTriState.msoFalse;
                }

            }


            // When there is no selection on the slide, disable the add/copy buttons on
            // customshapepane and syncpane respectively
            SyncPane syncPane = GetActivePane(typeof(SyncPane))?.Control as SyncPane;
            syncPane?.UpdateOnSelectionChange(sel);

            CustomShapePane customShapePane = GetActivePane(typeof(CustomShapePane))?.Control as CustomShapePane;
            customShapePane?.UpdateOnSelectionChange(sel);

            Ribbon.RefreshRibbonControl("AnimateInSlideButton");
            Ribbon.RefreshRibbonControl("DrillDownButton");
            Ribbon.RefreshRibbonControl("StepBackButton");
            Ribbon.RefreshRibbonControl("AddSpotlightButton");
            Ribbon.RefreshRibbonControl("AddZoomInButton");
            Ribbon.RefreshRibbonControl("AddZoomOutButton");
            Ribbon.RefreshRibbonControl("ZoomToAreaButton");
            Ribbon.RefreshRibbonControl("ReplaceWithClipboardButton");
            Ribbon.RefreshRibbonControl("PasteIntoGroupButton");
            Ribbon.RefreshRibbonControl("ConvertToTooltipButton");
            Ribbon.RefreshRibbonControl("CreateCalloutButton");
            Ribbon.RefreshRibbonControl("CreateTriggerButton");
            // To grey out the "HighlightText" button whenever non-text fragment or nothing has been selected
            Ribbon.RefreshRibbonControl("HighlightTextButton");
        }

        private void ThisAddInNewPresentation(PowerPoint.Presentation pres)
        {
            PowerPoint.DocumentWindow activeWindow = pres.Application.ActiveWindow;
            string tempName = pres.Name.GetHashCode().ToString(CultureInfo.InvariantCulture);

            _documentHashcodeMapper[activeWindow] = tempName;

            // Refresh ribbon to enable the menu buttons
            RefreshRibbonMenuButtons();
            // Initialise the "Maintain Tab Focus" and "Compress Images" checkbox
            Ribbon.InitialiseVisibilityCheckbox();
            Ribbon.InitialiseCompressImagesCheckbox();
        }

        // solve new un-modified unsave problem
        private void ThisAddInAfterNewPresentation(PowerPoint.Presentation pres)
        {
            //Access the BuiltInDocumentProperties so that the property storage does get created.
            object o = pres.BuiltInDocumentProperties;
            pres.Saved = Microsoft.Office.Core.MsoTriState.msoTrue;
        }

        private void ThisAddInPresentationOpen(PowerPoint.Presentation pres)
        {
            // Windows count could be zero if presentation is opened as preview of template slides
            if (pres.Application.Windows.Count > 0)
            {
                PowerPoint.DocumentWindow activeWindow = pres.Application.ActiveWindow;
                tempFolderName = pres.Name.GetHashCode().ToString(CultureInfo.InvariantCulture);

                // if we opened a new window, register the window with its name
                if (!_documentHashcodeMapper.ContainsKey(activeWindow))
                {
                    _documentHashcodeMapper[activeWindow] = tempFolderName;
                }

                // Refresh ribbon to enable the menu buttons if there are now at least one window
                RefreshRibbonMenuButtons();
                // Initialise the "Maintain Tab Focus" and "Compress Images" checkbox
                Ribbon.InitialiseVisibilityCheckbox();
                Ribbon.InitialiseCompressImagesCheckbox();
            }
        }

        private void ThisAddInPresentationClose(PowerPoint.Presentation pres)
        {
            Trace.TraceInformation("Closing " + pres.Name + "...");


            // We need to check if there is only one window AND the active window's presentation 
            // has the same name as the one we are closing because it is possible for background 
            // presentation (those without windows) to close and trigger this event as well.
            // We only want to shut down PPTLabs if we are closing the main presentation.
            if (Application.Windows.Count == 1 &&
                Application.ActiveWindow.Presentation.FullName == pres.FullName)
            {
                // If this current window we are closing is the last window, then PptLabs should terminate.
                _pptLabsShouldTerminate = true;
            }

            // special case: if we are closing 'ShapeGallery.pptx' or 'Sync Lab - Do not edit.pptx', no other action will be done
            if (pres.Name.Contains(ShapeGalleryPptxName) || pres.Name.Contains(SyncLabPptxName))
            {
                return;
            }

            if (_pptLabsShouldTerminate)
            {
                ShutDownSyncLab();
                ShutDownShapesLab();
                ShutDownPictureSlidesLab();
            }

            ShutDownRecorderPane();

            // find the document that holds the presentation with pres.Name
            // special case will be embedded slide. in this case pres.Windows return exception
            PowerPoint.DocumentWindow associatedWindow;

            try
            {
                Trace.TraceInformation("Total windows of closing presentation = " + pres.Windows.Count);
                Trace.TraceInformation("Windows are: ");

                foreach (PowerPoint.DocumentWindow window in pres.Windows)
                {
                    Trace.TraceInformation("\t" + window.Caption);
                }

                associatedWindow = pres.Windows[1];
            }
            catch (Exception)
            {
                Trace.TraceInformation("Closing presentation - " + pres.FullName + " - has no window.");
                return;
            }

            // for Functional Test to close presentation
            if (PowerPointLabsFT.IsFunctionalTestOn)
            {
                IntPtr handle = Native.FindWindow("PPTFrameClass", pres.Name + " - Microsoft PowerPoint");
                Native.SetForegroundWindow(handle);
                SendKeys.Send("N");
            }

            Trace.TraceInformation("Closing associated window...");
            CleanUp(associatedWindow);

            // Refresh ribbon to grey out the menu / buttons if there are no windows open
            RefreshRibbonMenuButtons();

        }

        private void RefreshRibbonMenuButtons()
        {
            Ribbon.RefreshRibbonControl(AnimationLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(ZoomLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(NarrationsLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(CaptionsLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(HighlightLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(EffectsLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(PositionsLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(ResizeLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(ColorsLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(SaveLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(SyncLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(ShapesLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(CropLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(PasteLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(TimerLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(AgendaLabText.RibbonMenuId);
            Ribbon.RefreshRibbonControl(PictureSlidesLabText.RibbonMenuId);
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            PPMouse.StopHook();
            PPKeyboard.StopHook();
            PPCopy.StopHook();
            // Event Handler unregistering taken care of in destructor
            UIThreadExecutor.TearDown();
            PPLClipboard.Instance.Teardown();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Exiting");
            Trace.Close();
            if (_ftChannel != null)
            {
                ChannelServices.UnregisterChannel(_ftChannel);
            }
        }

        #endregion

        # region Copy paste handlers

        private PowerPoint.DocumentWindow _copyFromWnd;
        private readonly Regex _shapeNamePattern = new Regex(@"^[^\[]\D+\s\d+$");
        private HashSet<String> _isShapeMatchedAlready;

        private void AfterPasteEventHandler(PowerPoint.Selection selection)
        {
            try
            {
                PowerPoint.Slide currentSlide = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                string pptName = Application.ActivePresentation.Name;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes
                    && currentSlide?.SlideID != _previousSlideForCopyEvent.SlideID
                    && pptName == _previousPptName)
                {
                    PowerPoint.ShapeRange pastedShapes = selection.ShapeRange;

                    List<string> nameListForPastedShapes = new List<string>();
                    Dictionary<string, string> nameDictForPastedShapes = new Dictionary<string, string>();
                    List<string> nameListForCopiedShapes = new List<string>();
                    List<PowerPoint.Shape> corruptedShapes = new List<PowerPoint.Shape>();

                    foreach (PowerPoint.Shape shape in _copiedShapes)
                    {
                        try
                        {
                            nameListForCopiedShapes.Add(shape.Name);
                        }
                        catch
                        {
                            //handling corrupted shapes
                            PowerPoint.Shape fixedShape = _previousSlideForCopyEvent.Shapes.SafeCopyPlaceholder(shape);
                            fixedShape.Left = shape.Left;
                            fixedShape.Top = shape.Top;
                            while (fixedShape.ZOrderPosition > shape.ZOrderPosition)
                            {
                                fixedShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                            }
                            corruptedShapes.Add(shape);
                            nameListForCopiedShapes.Add(fixedShape.Name);
                        }
                    }

                    foreach (PowerPoint.Shape shape in corruptedShapes)
                    {
                        shape.Delete();
                    }

                    _isShapeMatchedAlready = new HashSet<string>();

                    for (int i = 1; i <= pastedShapes.Count; i++)
                    {
                        PowerPoint.Shape shape = pastedShapes[i];
                        int matchedShapeIndex = FindMatchedShape(shape);
                        string uniqueName = Guid.NewGuid().ToString();
                        nameDictForPastedShapes[uniqueName] = nameListForCopiedShapes[matchedShapeIndex];
                        shape.Name = uniqueName;
                        nameListForPastedShapes.Add(shape.Name);
                    }
                    //Re-select pasted shapes
                    PowerPoint.ShapeRange range = currentSlide.Shapes.Range(nameListForPastedShapes.ToArray());
                    foreach (PowerPoint.Shape shape in range)
                    {
                        shape.Name = nameDictForPastedShapes[shape.Name];
                    }
                    range.Select();
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    List<PowerPoint.Slide> pastedSlides = selection.SlideRange.Cast<PowerPoint.Slide>().OrderBy(x => x.SlideIndex).ToList();

                    for (int i = 0; i < pastedSlides.Count; i++)
                    {
                        if (AgendaLab.AgendaSlide.IsReferenceslide(_copiedSlides[i]))
                        {
                            pastedSlides[i].Name = _copiedSlides[i].Name;
                            pastedSlides[i].Design = _copiedSlides[i].Design;
                        }
                    }
                }
            }
            catch
            {
                //TODO: log in ThisAddIn.cs
            }
        }

        private int FindMatchedShape(PowerPoint.Shape shape)
        {
            //Strong matching:
            for (int i = 0; i < _copiedShapes.Count; i++)
            {
                if (IsSimilarShape(shape, _copiedShapes[i])
                    && IsSimilarName(shape.Name, _copiedShapes[i].Name)
                    && Math.Abs(shape.Left - _copiedShapes[i].Left) < float.Epsilon
                    && Math.Abs(shape.Height - _copiedShapes[i].Height) < float.Epsilon
                    && !_isShapeMatchedAlready.Contains(_copiedShapes[i].Id.ToString(CultureInfo.InvariantCulture)))
                {
                    _isShapeMatchedAlready.Add(_copiedShapes[i].Id.ToString(CultureInfo.InvariantCulture));

                    return i;
                }
            }
            //Blur matching:
            for (int i = 0; i < _copiedShapes.Count; i++)
            {
                if (IsSimilarShape(shape, _copiedShapes[i])
                    && IsSimilarName(shape.Name, _copiedShapes[i].Name)
                    && !_isShapeMatchedAlready.Contains(_copiedShapes[i].Id.ToString(CultureInfo.InvariantCulture)))
                {
                    _isShapeMatchedAlready.Add(_copiedShapes[i].Id.ToString(CultureInfo.InvariantCulture));

                    return i;
                }
            }
            return -1;
        }

        private bool IsSimilarShape(PowerPoint.Shape shape, PowerPoint.Shape shape2)
        {
            return Math.Abs(shape.Width - shape2.Width) < float.Epsilon
                   && Math.Abs(shape.Height - shape2.Height) < float.Epsilon
                   && shape.Type == shape2.Type
                   && (shape.Type != Office.MsoShapeType.msoAutoShape
                       || shape.AutoShapeType == shape2.AutoShapeType);
        }

        /// <summary>
        /// Similar name defi:
        /// 1. if they're not default shape name, they must be the exact same
        /// 2. if they're default shape name, the shape type in the name must be the exact same
        /// 3. otherwise not similar
        /// </summary>
        /// <param name="name1"></param>
        /// <param name="name2"></param>
        /// <returns></returns>
        private bool IsSimilarName(string name1, string name2)
        {
            //remove enclosing brackets for name2
            Regex nameEnclosedInBrackets = new Regex(@"^\[\D+\s\d+\]$");
            if (!nameEnclosedInBrackets.IsMatch(name1)
                && nameEnclosedInBrackets.IsMatch(name2)
                && name2.Length > 2)
            {
                name2 = name2.Substring(1, name2.Length - 2);
            }

            if (!_shapeNamePattern.IsMatch(name1)
                && !_shapeNamePattern.IsMatch(name2))
            {
                return name1.Equals(name2);
            }

            if (_shapeNamePattern.IsMatch(name1)
                && _shapeNamePattern.IsMatch(name2))
            {
                Regex shapeTypeInName = new Regex(@"^[^\[]\D+\s(?=\d+$)");
                string shapeTypeForName1 = shapeTypeInName.Match(name1).ToString();
                string shapeTypeForName2 = shapeTypeInName.Match(name2).ToString();
                return shapeTypeForName1.Equals(shapeTypeForName2);
            }
            return false;
        }

        private void AfterPasteRecorderEventHandler(PowerPoint.Selection selection)
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                // invalid paste event triggered because of system message loss
                if (_copiedSlides.Count < 1)
                {
                    return;
                }

                // if we copied from a presentation without recorder pane or pasted to a
                // presentation without recorder pane, paste event will not be entertained
                if (!_documentPaneMapper.ContainsKey(_copyFromWnd) ||
                    _documentPaneMapper[_copyFromWnd] == null ||
                    GetActivePane(typeof(RecorderTaskPane)) == null)
                {
                    return;
                }

                RecorderTaskPane copyFromRecorderPane =
                    GetPaneFromWindow(typeof(RecorderTaskPane), _copyFromWnd).Control as RecorderTaskPane;
                RecorderTaskPane activeRecorderPane = GetActivePane(typeof(RecorderTaskPane)).Control as RecorderTaskPane;

                if (activeRecorderPane == null ||
                    copyFromRecorderPane == null)
                {
                    return;
                }

                PowerPoint.SlideRange slideRange = selection.SlideRange;
                int oriSlide = 0;

                foreach (object sld in slideRange)
                {
                    PowerPointSlide oldSlide = PowerPointSlide.FromSlideFactory(_copiedSlides[oriSlide]);
                    PowerPointSlide newSlide = PowerPointSlide.FromSlideFactory(sld as PowerPoint.Slide);

                    activeRecorderPane.PasteSlideAudioAndScript(newSlide,
                                                                copyFromRecorderPane.CopySlideAudioAndScript(oldSlide));

                    oriSlide++;
                }

                // update the lists when all done
                UpdateRecorderPane(slideRange.Count, slideRange[1].SlideID);
            }
        }

        private void AfterCopyEventHandler(PowerPoint.Selection selection)
        {
            try
            {
                _copyFromWnd = Application.ActiveWindow;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    _copiedSlides.Clear();

                    foreach (object sld in selection.SlideRange)
                    {
                        PowerPoint.Slide slide = sld as PowerPoint.Slide;

                        _copiedSlides.Add(slide);
                    }

                    _copiedSlides.Sort((x, y) => (x.SlideIndex - y.SlideIndex));
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    _copiedShapes.Clear();
                    _previousSlideForCopyEvent = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                    _previousPptName = Application.ActivePresentation.Name;
                    foreach (object sh in selection.ShapeRange)
                    {
                        PowerPoint.Shape shape = sh as PowerPoint.Shape;
                        _copiedShapes.Add(shape);
                    }

                    _copiedShapes.Sort((x, y) => (x.Id - y.Id));
                }
                Ribbon.RefreshRibbonControl("PasteToFillSlideButton");
                Ribbon.RefreshRibbonControl("PasteToFitSlideButton");
                Ribbon.RefreshRibbonControl("PasteAtOriginalPositionButton");
                Ribbon.RefreshRibbonControl("ReplaceWithClipboardButton");
                Ribbon.RefreshRibbonControl("PasteIntoGroupButton");
            }
            catch
            {
                //TODO: log in ThisAddIn.cs
            }
        }
        # endregion

        #region Tab Activate

        private void SetupTabActivateHandler()
        {
            _tabActivate += TabActivateEventHandler;
        }

        private Native.WinEventDelegate _tabActivate;

        private IntPtr _eventHook = IntPtr.Zero;

        //This handler is used to check, whether Home tab is enabled or not
        //After Shortcut (Alt + H + O) is sent to PowerPoint by method OpenPropertyWindowForOffice10,
        //if unsuccessful (Home tab is not enabled), EVENT_SYSTEM_MENUEND will be received
        //if successful   (Property window is open), EVENT_OBJECT_CREATE will be received
        //To check the events occurred, use AccEvent32.exe
        //Refer to MSAA - Event Constants:
        //http://msdn.microsoft.com/en-us/library/windows/desktop/dd318066(v=vs.85).aspx
        void TabActivateEventHandler(IntPtr hook, uint eventType,
        IntPtr hwnd, int idObject, int child, uint thread, uint time)
        {
            if (eventType == (uint)Native.Event.EVENT_SYSTEM_MENUEND
                || eventType == (uint)Native.Event.EVENT_OBJECT_CREATE)
            {
                Native.UnhookWinEvent(_eventHook);
                _eventHook = IntPtr.Zero;
            }
            if (eventType == (uint)Native.Event.EVENT_SYSTEM_MENUEND)
            {
                MessageBox.Show(CommonText.ErrorTabActivate, CommonText.ErrorTabActivateTitle);
            }
        }

        #endregion

        #region Double Click to Open Property Window
        private const string ShortcutAltHO = "%h%o";

        private const int CommandOpenBackgroundFormat = 0x8F;

        private bool _isInSlideShow;

        private void SetupAfterCopyPasteHandler()
        {
            PPCopy.AfterCopy += AfterCopyEventHandler;
            PPCopy.AfterPaste += AfterPasteRecorderEventHandler;
            PPCopy.AfterPaste += AfterPasteEventHandler;
        }

        private readonly List<PowerPoint.Shape> _copiedShapes = new List<PowerPoint.Shape>();
        private readonly List<PowerPoint.Slide> _copiedSlides = new List<PowerPoint.Slide>();
        private PowerPoint.Slide _previousSlideForCopyEvent;
        private string _previousPptName;

        private void SetupDoubleClickHandler()
        {
            PPMouse.DoubleClick += DoubleClickEventHandler;
        }

        private void DoubleClickEventHandler(PowerPoint.Selection selection)
        {
            try
            {
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    TrySelectTransparentShape();
                }

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    if (IsApplicationVersion2010())
                    {
                        OpenPropertyWindowForOffice10();
                    }
                    else
                    {
                        OpenPropertyWindowForOffice13OrHigher(selection);
                    }
                }
            }
            catch (COMException e)
            {
                string logText = "DoubleClickEventHandler" + ": " + e.Message + ": " + e.StackTrace;
                Trace.TraceError(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            }
        }

        private void TrySelectTransparentShape()
        {
            if (PowerPointCurrentPresentationInfo.CurrentSlide == null)
            {
                return;
            }

            PowerPoint.Shape overlappingShape = null;
            int overlappingShapeZIndex = -1;

            PowerPoint.Shapes shapesInCurrentSlide = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;
            foreach (PowerPoint.Shape shape in shapesInCurrentSlide)
            {
                if (IsMouseWithinShape(shape)
                    && shape.ZOrderPosition > overlappingShapeZIndex)
                {
                    overlappingShape = shape;
                    overlappingShapeZIndex = shape.ZOrderPosition;
                }
            }
            if (overlappingShape?.Visible == Office.MsoTriState.msoTrue)
            {
                overlappingShape.Select();
            }
        }

        private bool IsMouseWithinShape(PowerPoint.Shape sh)
        {
            float x = Cursor.Position.X;
            float y = Cursor.Position.Y;
            int left = Application.ActiveWindow.PointsToScreenPixelsX(sh.Left);
            int top = Application.ActiveWindow.PointsToScreenPixelsY(sh.Top);
            int right = Application.ActiveWindow.PointsToScreenPixelsX(sh.Left + sh.Width);
            int bottom = Application.ActiveWindow.PointsToScreenPixelsY(sh.Top + sh.Height);
            return x > left
                && x < right
                && y > top
                && y < bottom;
        }

        //For office 2013 or Higher version:
        //Open Background Format window, then selecting the shape will
        //convert the window to Property window
        private void OpenPropertyWindowForOffice13OrHigher(PowerPoint.Selection selection)
        {
            if (!_isInSlideShow)
            {
                PowerPoint.ShapeRange selectedShapes = selection.ShapeRange;
                Native.SendMessage(
                    Process.GetCurrentProcess().MainWindowHandle,
                    (uint)Native.Message.WM_COMMAND,
                    new IntPtr(CommandOpenBackgroundFormat),
                    IntPtr.Zero);
                selectedShapes.Select();
            }
        }

        //For office 2010 (in office 2013, this method has bad user exp)
        //Use hotkey (Alt - H - O) to activate Property window
        private void OpenPropertyWindowForOffice10()
        {
            try
            {
                if (!_isInSlideShow)
                {
                    if (_eventHook == IntPtr.Zero)
                    {
                        //Check whether Home tab is enabled or not
                        _eventHook = Native.SetWinEventHook(
                            (uint)Native.Event.EVENT_SYSTEM_MENUEND,
                            (uint)Native.Event.EVENT_OBJECT_CREATE,
                            IntPtr.Zero,
                            _tabActivate,
                            (uint)Process.GetCurrentProcess().Id,
                            0,
                            0);
                    }
                    SendKeys.Send(ShortcutAltHO);
                }
            }
            catch (InvalidOperationException)
            {
                // ignore exception
            }
        }
        #endregion

        private void SetupFunctionalTestChannels()
        {
            _ftChannel = new IpcChannel("PowerPointLabsFT");
            ChannelServices.RegisterChannel(_ftChannel, false);
            RemotingConfiguration.RegisterWellKnownServiceType(typeof(PowerPointLabsFT),
                "PowerPointLabsFT", WellKnownObjectMode.Singleton);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddInStartup;
            Shutdown += ThisAddInShutdown;
        }

        #endregion
    }
}
