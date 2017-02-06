using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Tools;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.FunctionalTestInterface.Impl;
using PowerPointLabs.FunctionalTestInterface.Impl.Controller;
using PowerPointLabs.Models;
using PowerPointLabs.PositionsLab;
using PowerPointLabs.ResizeLab;
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
        private const string AppLogName = "PowerPointLabs.log"; 
        private const string SlideXmlSearchPattern = @"slide(\d+)\.xml";
        private const string TempFolderNamePrefix = @"\PowerPointLabs Temp\";
        private const string ShapeGalleryPptxName = "ShapeGallery";
        private const string TempZipName = "tempZip.zip";

        private string _deactivatedPresFullName;

        private bool _isClosing;

        private bool isResizePaneVisible;

        private readonly Dictionary<PowerPoint.DocumentWindow,
            List<CustomTaskPane>> _documentPaneMapper = new Dictionary<PowerPoint.DocumentWindow,
                List<CustomTaskPane>>();

        private readonly Dictionary<PowerPoint.DocumentWindow,
            string> _documentHashcodeMapper = new Dictionary<PowerPoint.DocumentWindow,
                string>();

        internal ShapesLabConfig ShapesLabConfigs;

        internal PowerPointShapeGalleryPresentation ShapePresentation;

        public readonly string OfficeVersion2013 = "15.0";
        public readonly string OfficeVersion2010 = "14.0";

        public static string AppDataFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PowerPointLabs");

        public Ribbon1 Ribbon;

        /// <summary>
        /// The channel for .NET Remoting calls.
        /// </summary>
        private IChannel _ftChannel;

        private void SetupFunctionalTestChannels()
        {
            _ftChannel = new IpcChannel("PowerPointLabsFT");
            ChannelServices.RegisterChannel(_ftChannel, false);
            RemotingConfiguration.RegisterWellKnownServiceType(typeof(PowerPointLabsFT),
                "PowerPointLabsFT", WellKnownObjectMode.Singleton);
        }

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

            // According to MSDN, when more than 1 event are triggered, callback's invoking sequence
            // follows the defining order. I.e. the earlier you defined, the earlier it will be
            // executed.

            // Here, we want the priority to be: Application action > Window action > Slide action

            // Priority High: Application Actions
            ((PowerPoint.EApplication_Event) Application).NewPresentation += ThisAddInNewPresentation;
            Application.AfterNewPresentation += ThisAddInAfterNewPresentation;
            Application.PresentationOpen += ThisAddInPrensentationOpen;
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
            Trace.TraceInformation(pres.Name + " terminating...");
            Trace.TraceInformation(string.Format("Is Closing = {0}, Count = {1}", _isClosing,
                Application.Presentations.Count));

            _deactivatedPresFullName = pres.FullName;

            // in this case, we are closing the last client presentation, therefore we
            // we can close the shape gallery
            if (_isClosing &&
                Application.Presentations.Count == 2 &&
                ShapePresentation != null &&
                ShapePresentation.Opened)
            {
                if (string.IsNullOrEmpty(ShapesLabConfigs.DefaultCategory))
                {
                    ShapesLabConfigs.DefaultCategory = ShapePresentation.Categories[0];
                }

                ShapePresentation.Close();
                Trace.TraceInformation("Shape Gallery terminated.");
            }
        }

        private void ThisAddInApplicationOnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            if (pres != null)
            {
                Ribbon.EmbedAudioVisible = !pres.Name.EndsWith(".ppt");

                var customShape = GetActiveControl(typeof(CustomShapePane)) as CustomShapePane;

                // make sure ShapeGallery's default category is consistent with current presentation
                if (customShape != null)
                {
                    var currentCategory = customShape.CurrentCategory;
                    ShapePresentation.DefaultCategory = currentCategory;
                }

                _isClosing = false;
            }
        }

        private void ThisAddInSlideSelectionChanged(PowerPoint.SlideRange sldRange)
        {
            // TODO: doing range sweep to check these var may affect performance, consider initializing these
            // TODO: variables only at program starts
            Ribbon.RemoveCaptionsEnabled = SlidesInRangeHaveCaptions(sldRange);
            Ribbon.RemoveAudioEnabled = SlidesInRangeHaveAudio(sldRange);

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
            }
            else
            {
                UpdateRecorderPane(sldRange.Count, -1);
            }

            // in case the recorder is on event
            BreakRecorderEvents();

            // ribbon function init
            Ribbon.AddAutoMotionEnabled = true;
            Ribbon.ReloadAutoMotionEnabled = true;
            Ribbon.ReloadSpotlight = true;
            Ribbon.HighlightBulletsEnabled = true;

            if (sldRange.Count != 1)
            {
                Ribbon.AddAutoMotionEnabled = false;
                Ribbon.ReloadAutoMotionEnabled = false;
                Ribbon.ReloadSpotlight = false;
                Ribbon.HighlightBulletsEnabled = false;
            }
            else
            {
                PowerPoint.Slide tmp = sldRange[1];
                PowerPoint.Presentation presentation = PowerPointPresentation.Current.Presentation;
                int slideIndex = tmp.SlideIndex;
                PowerPoint.Slide next = tmp;
                PowerPoint.Slide prev = tmp;

                if (slideIndex < presentation.Slides.Count)
                    next = presentation.Slides[slideIndex + 1];
                if (slideIndex > 1)
                    prev = presentation.Slides[slideIndex - 1];
                if (!((tmp.Name.StartsWith("PPSlideAnimated"))
                      || ((tmp.Name.StartsWith("PPSlideStart"))
                          && (next.Name.StartsWith("PPSlideAnimated")))
                      || ((tmp.Name.StartsWith("PPSlideEnd"))
                          && (prev.Name.StartsWith("PPSlideAnimated")))
                      || ((tmp.Name.StartsWith("PPSlideMulti"))
                          && ((prev.Name.StartsWith("PPSlideAnimated"))
                              || (next.Name.StartsWith("PPSlideAnimated"))))))
                    Ribbon.ReloadAutoMotionEnabled = false;
                if (!(tmp.Name.Contains("PPTLabsSpotlight")))
                    Ribbon.ReloadSpotlight = false;
            }

            Ribbon.RefreshRibbonControl("AddAnimationButton");
            Ribbon.RefreshRibbonControl("ReloadButton");
            Ribbon.RefreshRibbonControl("ReloadSpotlightButton");
            Ribbon.RefreshRibbonControl("HighlightBulletsTextButton");
            Ribbon.RefreshRibbonControl("HighlightBulletsBackgroundButton");
            Ribbon.RefreshRibbonControl("RemoveCaptionsButton");
            Ribbon.RefreshRibbonControl("RemoveAudioButton");
        }

        private void ThisAddInSelectionChanged(PowerPoint.Selection sel)
        {
            Ribbon.SpotlightEnabled = false;
            Ribbon.InSlideEnabled = false;
            Ribbon.ZoomButtonEnabled = false;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape sh = sel.ShapeRange[1];
                if (sh.Type == Office.MsoShapeType.msoAutoShape || sh.Type == Office.MsoShapeType.msoFreeform ||
                    sh.Type == Office.MsoShapeType.msoTextBox || sh.Type == Office.MsoShapeType.msoPlaceholder
                    || sh.Type == Office.MsoShapeType.msoCallout || sh.Type == Office.MsoShapeType.msoInk ||
                    sh.Type == Office.MsoShapeType.msoGroup)
                {
                    Ribbon.SpotlightEnabled = true;
                }
                if ((sh.Type == Office.MsoShapeType.msoAutoShape &&
                     sh.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle) ||
                    sh.Type == Office.MsoShapeType.msoPicture)
                {
                    Ribbon.ZoomButtonEnabled = true;
                }
                if (sel.ShapeRange.Count > 1)
                {
                    foreach (PowerPoint.Shape tempShape in sel.ShapeRange)
                    {
                        if (sh.Type == tempShape.Type)
                        {
                            Ribbon.InSlideEnabled = true;
                            Ribbon.ZoomButtonEnabled = true;
                        }
                        if (sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType != tempShape.AutoShapeType)
                        {
                            Ribbon.InSlideEnabled = false;
                            Ribbon.ZoomButtonEnabled = false;
                            break;
                        }
                    }
                }

                if (isResizePaneVisible)
                {
                    sel.ShapeRange.LockAspectRatio = ResizeLabPaneWPF.IsAspectRatioLocked
                        ? Office.MsoTriState.msoTrue
                        : Office.MsoTriState.msoFalse;
                }

            }

            Ribbon.RefreshRibbonControl("AddSpotlightButton");
            Ribbon.RefreshRibbonControl("InSlideAnimateButton");
            Ribbon.RefreshRibbonControl("AddZoomInButton");
            Ribbon.RefreshRibbonControl("AddZoomOutButton");
            Ribbon.RefreshRibbonControl("ZoomToAreaButton");
        }

        private void ThisAddInNewPresentation(PowerPoint.Presentation pres)
        {
            var activeWindow = pres.Application.ActiveWindow;
            var tempName = pres.Name.GetHashCode().ToString(CultureInfo.InvariantCulture);

            _documentHashcodeMapper[activeWindow] = tempName;
        }

        // solve new un-modified unsave problem
        private void ThisAddInAfterNewPresentation(PowerPoint.Presentation pres)
        {
            //Access the BuiltInDocumentProperties so that the property storage does get created.
            object o = pres.BuiltInDocumentProperties;
            pres.Saved = Microsoft.Office.Core.MsoTriState.msoTrue;
        }

        private void ThisAddInPrensentationOpen(PowerPoint.Presentation pres)
        {
            var activeWindow = pres.Application.ActiveWindow;
            var tempName = pres.Name.GetHashCode().ToString(CultureInfo.InvariantCulture);

            // if we opened a new window, register the window with its name
            if (!_documentHashcodeMapper.ContainsKey(activeWindow))
            {
                _documentHashcodeMapper[activeWindow] = tempName;
            }
        }

        private void ThisAddInPresentationClose(PowerPoint.Presentation pres)
        {
            Trace.TraceInformation("Closing " + pres.Name);

            if (Application.Version == OfficeVersion2010 &&
                _deactivatedPresFullName == pres.FullName &&
                Application.Presentations.Count == 2 &&
                ShapePresentation != null &&
                ShapePresentation.Opened)
            {
                ShapePresentation.Close();
            }

            // special case: if we are closing ShapeGallery.pptx, no other action will be done
            if (pres.Name.Contains(ShapeGalleryPptxName))
            {
                return;
            }

            ShutDownColorPane();
            ShutDownRecorderPane();
            ShutDownImageSearchPane();

            // find the document that holds the presentation with pres.Name
            // special case will be embedded slide. in this case pres.Windows return exception
            PowerPoint.DocumentWindow associatedWindow;

            try
            {
                Trace.TraceInformation("Total Windows at Close Stage " + pres.Windows.Count);
                Trace.TraceInformation("Windows are: ");

                foreach (PowerPoint.DocumentWindow window in pres.Windows)
                {
                    Trace.TraceInformation(window.Presentation.Name);
                }

                associatedWindow = pres.Windows[1];
            }
            catch (Exception)
            {
                return;
            }

            // for Functional Test to close presentation
            if (PowerPointLabsFT.IsFunctionalTestOn)
            {
                var handle = Native.FindWindow("PPTFrameClass", pres.Name + " - Microsoft PowerPoint");
                Native.SetForegroundWindow(handle);
                SendKeys.Send("N");
            }

            Trace.TraceInformation("Closing associated window...");
            CleanUp(associatedWindow);
        }

        private void ShutDownImageSearchPane()
        {
            var pictureSlidesLabWindow = Globals.ThisAddIn.Ribbon.PictureSlidesLabWindow;
            if (pictureSlidesLabWindow != null && pictureSlidesLabWindow.IsOpen && Application.Presentations.Count == 2)
            {
                pictureSlidesLabWindow.Close();
            }
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            PPMouse.StopHook();
            PPKeyboard.StopHook();
            PPCopy.StopHook();
            PositionsPaneWpf.ClearAllEventHandlers();
            UIThreadExecutor.TearDown();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Exiting");
            Trace.Close();
            if (_ftChannel != null)
            {
                ChannelServices.UnregisterChannel(_ftChannel);
            }
        }

        # endregion

        # region API

        public Control GetActiveControl(Type type)
        {
            var taskPane = GetActivePane(type);

            return taskPane == null ? null : taskPane.Control;
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
            var taskPane = GetPaneFromWindow(typeof(CustomShapePane), window);

            return taskPane == null ? null : taskPane.Control;
        }

        public CustomTaskPane GetPaneFromWindow(Type type, PowerPoint.DocumentWindow window)
        {
            if (!_documentPaneMapper.ContainsKey(window))
            {
                return null;
            }

            var panes = _documentPaneMapper[window];

            foreach (var pane in panes)
            {
                try
                {
                    var control = pane.Control;

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
            if (ShapePresentation != null && ShapePresentation.Opened) return;

            var shapeRootFolderPath = ShapesLabConfigs.ShapeRootFolder;

            ShapePresentation =
                new PowerPointShapeGalleryPresentation(shapeRootFolderPath, ShapeGalleryPptxName);

            if (!ShapePresentation.Open(withWindow: false, focus: false) &&
                !ShapePresentation.Opened)
            {
                // if the presentation gets some error during opening, and the error could not
                // be resolved by consistency check, prompt the user about the error
                MessageBox.Show(TextCollection.ShapeGalleryInitErrorMsg);
                return;
            }

            if (ShapePresentation.HasCategory(ShapesLabConfigs.DefaultCategory))
            {
                ShapePresentation.DefaultCategory = ShapesLabConfigs.DefaultCategory;

                return;
            }

            // if we do not have the default category, create and add it to ShapeGallery
            ShapePresentation.AddCategory(ShapesLabConfigs.DefaultCategory);
            ShapePresentation.Save();
        }

        public void InitializeShapesLabConfig()
        {
            // if ShapesLabConfig has already been intialized, do nothing
            if (ShapesLabConfigs != null) return;

            ShapesLabConfigs = new ShapesLabConfig(AppDataFolder);

            // create a directory under specified location if the location does not exist
            if (!Directory.Exists(ShapesLabConfigs.ShapeRootFolder))
            {
                Directory.CreateDirectory(ShapesLabConfigs.ShapeRootFolder);
            }
        }

        public void PrepareMediaFiles(PowerPoint.Presentation pres, string tempPath)
        {
            var presFullName = pres.FullName;
            var presName = pres.Name;

            // in case of embedded slides, we need to regulate the file name and full name
            RegulatePresentationName(pres, tempPath, ref presName, ref presFullName);

            try
            {
                if (IsEmptyFile(presFullName))
                {
                    return;
                }

                var zipFullPath = tempPath + TempZipName;

                // before we do everything, check if there's an undelete old zip file
                // due to some error
                try
                {
                    FileDir.DeleteFile(zipFullPath);
                    FileDir.CopyFile(presFullName, zipFullPath);
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog(TextCollection.AccessTempFolderErrorMsg, string.Empty, e);
                }

                ExtractMediaFiles(zipFullPath, tempPath);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog(TextCollection.PrepareMediaErrorMsg, "Files cannot be linked.", e);
            }
        }

        public string PrepareTempFolder(PowerPoint.Presentation pres)
        {
            var presName = pres.Name;
            var presFullName = pres.FullName;

            // here presFullName makes no use, just to fit in the signature
            RegulatePresentationName(pres, null, ref presName, ref presFullName);

            var tempPath = GetPresentationTempFolder(presName);

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
                ErrorDialogWrapper.ShowDialog(TextCollection.CreatTempFolderErrorMsg, string.Empty, e);
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

            var activeWindow = presentation.Application.ActiveWindow;

            RegisterTaskPane(new ResizeLabPane(), TextCollection.ResizeLabsTaskPaneTitle, activeWindow, 
                ResizeTaskPaneVisibleValueChangedEventHandler, null);
        }

        public void RegisterRecorderPane(PowerPoint.DocumentWindow activeWindow, string tempFullPath)
        {
            if (GetActivePane(typeof(RecorderTaskPane)) != null)
            {
                return;
            }

            RegisterTaskPane(new RecorderTaskPane(tempFullPath), TextCollection.RecManagementPanelTitle, activeWindow,
                TaskPaneVisibleValueChangedEventHandler, null);
        }

        public void RegisterColorPane(PowerPoint.Presentation presentation)
        {
            if (GetActivePane(typeof(ColorPane)) != null)
            {
                return;
            }

            var activeWindow = presentation.Application.ActiveWindow;

            RegisterTaskPane(new ColorPane(), TextCollection.ColorsLabTaskPanelTitle, activeWindow, null, null);
        }

        public void RegisterDrawingsPane(PowerPoint.Presentation presentation)
        {
            if (GetActivePane(typeof(DrawingsPane)) != null)
            {
                return;
            }

            var activeWindow = presentation.Application.ActiveWindow;

            RegisterTaskPane(new DrawingsPane(), TextCollection.DrawingsLabTaskPanelTitle, activeWindow, null, null);
        }

        public void RegisterShapesLabPane(PowerPoint.Presentation presentation)
        {
            if (GetActivePane(typeof(CustomShapePane)) != null)
            {
                return;
            }

            var activeWindow = presentation.Application.ActiveWindow;

            RegisterTaskPane(
                new CustomShapePane(ShapesLabConfigs.ShapeRootFolder, ShapesLabConfigs.DefaultCategory),
                TextCollection.ShapesLabTaskPanelTitle, activeWindow, null, null);
        }

        public void SyncShapeAdd(string shapeName, string shapeFullName, string category)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow) continue;

                var shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl != null &&
                    shapePaneControl.CurrentCategory == category)
                {
                    shapePaneControl.AddCustomShape(shapeName, shapeFullName, false);
                }
            }
        }

        public void SyncShapeRemove(string shapeName, string category)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow) continue;

                var shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl != null &&
                    shapePaneControl.CurrentCategory == category)
                {
                    shapePaneControl.RemoveCustomShape(shapeName);
                }
            }
        }

        public void SyncShapeRename(string shapeOldName, string shapeNewName, string category)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow) continue;

                var shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl != null &&
                    shapePaneControl.CurrentCategory == category)
                {
                    shapePaneControl.RenameCustomShape(shapeOldName, shapeNewName);
                }
            }
        }

        public bool VerifyOnLocal(PowerPoint.Presentation pres)
        {
            var invalidPathRegex = new Regex("^[hH]ttps?:");

            return !invalidPathRegex.IsMatch(pres.Path);
        }

        public bool VerifyVersion(PowerPoint.Presentation pres)
        {
            return !pres.Name.EndsWith(".ppt");
        }

        # endregion

        # region Helper Functions

        private void SetupLogger()
        {
            // Check if folder exists and if not, create it
            if (!Directory.Exists(AppDataFolder))
                Directory.CreateDirectory(AppDataFolder);

            var fileName = DateTime.Now.ToString("yyyy-MM-dd") + AppLogName;
            var logPath = Path.Combine(AppDataFolder, fileName);

            Trace.AutoFlush = true;
            Trace.Listeners.Add(new TextWriterTraceListener(logPath));
        }

        private void ShutDownRecorderPane()
        {
            var recorder = GetActiveControl(typeof(RecorderTaskPane)) as RecorderTaskPane;

            if (recorder != null &&
                recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }
        }

        private void ShutDownColorPane()
        {
            var colorPane = GetActivePane(typeof(ColorPane));

            if (colorPane == null) return;

            var colorLabs = colorPane.Control as ColorPane;
            if (colorLabs != null) colorLabs.SaveDefaultColorPaneThemeColors();
        }

        public CustomTaskPane RegisterTaskPane(UserControl control, string title, PowerPoint.DocumentWindow wnd,
            EventHandler visibleChangeEventHandler = null,
            EventHandler dockPositionChangeEventHandler = null)
        {
            var loadingDialog = new LoadingDialog();
            loadingDialog.Show();
            loadingDialog.Refresh();

            // note down the control's width
            var width = control.Width;

            // register the user control to the CustomTaskPanes collection and set it as
            // current active task pane;
            var taskPane = CustomTaskPanes.Add(control, title, wnd);

            // task pane UI setup
            taskPane.Visible = false;
            taskPane.Width = width + 20;

            // map the current window with the task pane
            if (!_documentPaneMapper.ContainsKey(wnd))
            {
                _documentPaneMapper[wnd] = new List<CustomTaskPane>();
            }

            _documentPaneMapper[wnd].Add(taskPane);

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

            loadingDialog.Dispose();
            return taskPane;
        }

        private void RemoveTaskPanes(PowerPoint.DocumentWindow activeWindow)
        {
            if (!_documentPaneMapper.ContainsKey(activeWindow))
            {
                return;
            }

            var activePanes = _documentPaneMapper[activeWindow];
            foreach (var pane in activePanes)
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

            var activePanes = _documentPaneMapper[window];
            for (var i = activePanes.Count - 1; i >= 0; i--)
            {
                var pane = activePanes[i];
                if (pane.Control.GetType() != paneType) continue;
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
            var recorderPane = GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            // trigger close form event when closing hide the pane
            if (recorder != null && !recorderPane.Visible)
            {
                recorder.RecorderPaneClosing();
                // remove recorder pane and force it to reload when next time open
                RemoveTaskPane(Application.ActiveWindow, typeof(RecorderTaskPane));
            }
        }

        private void ResizeTaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            var resizePane = GetActivePane(typeof(ResizeLabPane));

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

            var recorder = GetActiveControl(typeof(RecorderTaskPane)) as RecorderTaskPane;

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
                var slides = PowerPointPresentation.Current.Slides.ToList();

                for (var i = 0; i < recorder.AudioBuffer.Count; i++)
                {
                    if (recorder.AudioBuffer[i].Count == 0) continue;

                    foreach (var audio in recorder.AudioBuffer[i])
                    {
                        audio.Item1.EmbedOnSlide(slides[i], audio.Item2);

                        if (Globals.ThisAddIn.Ribbon.RemoveAudioEnabled) continue;

                        Globals.ThisAddIn.Ribbon.RemoveAudioEnabled = true;
                        Globals.ThisAddIn.Ribbon.RefreshRibbonControl("RemoveAudioButton");
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

            var fileInfo = new FileInfo(filePath);

            return fileInfo.Length == 0;
        }

        private void UpdateRecorderPane(int count, int id)
        {
            var recorderPane = GetActivePane(typeof(RecorderTaskPane));

            // if there's no active pane associated with the current window, return
            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

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

        private string GetPresentationTempFolder(string presName)
        {
            var tempName = presName.GetHashCode().ToString(CultureInfo.InvariantCulture);
            var tempPath = Path.GetTempPath() + TempFolderNamePrefix + tempName + @"\";

            return tempPath;
        }

        private void CleanUp(PowerPoint.DocumentWindow associatedWindow)
        {
            _isClosing = true;

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
                var zip = ZipStorer.Open(zipFullPath, FileAccess.Read);
                var dir = zip.ReadCentralDir();

                var regex = new Regex(SlideXmlSearchPattern);

                foreach (var entry in dir)
                {
                    var name = Path.GetFileName(entry.FilenameInZip);

                    if (name == null) continue;

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
                ErrorDialogWrapper.ShowDialog(TextCollection.ExtraErrorMsg, "Archived files cannot be retrieved.", e);
            }
        }

        private void BreakRecorderEvents()
        {
            var recorder = GetActiveControl(typeof(RecorderTaskPane)) as RecorderTaskPane;

            if (recorder != null &&
                recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }
        }
        # endregion

        # region Copy paste handlers

        private PowerPoint.DocumentWindow _copyFromWnd;
        private readonly Regex _shapeNamePattern = new Regex(@"^[^\[]\D+\s\d+$");
        private HashSet<String> _isShapeMatchedAlready;

        private void AfterPasteEventHandler(PowerPoint.Selection selection)
        {
            try
            {
                var currentSlide = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                var pptName = Application.ActivePresentation.Name;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes
                    && currentSlide != null
                    && currentSlide.SlideID != _previousSlideForCopyEvent.SlideID
                    && pptName == _previousPptName)
                {
                    PowerPoint.ShapeRange pastedShapes = selection.ShapeRange;

                    var nameListForPastedShapes = new List<string>();
                    var nameDictForPastedShapes = new Dictionary<string, string>();
                    var nameListForCopiedShapes = new List<string>();
                    var corruptedShapes = new List<PowerPoint.Shape>();

                    foreach (var shape in _copiedShapes)
                    {
                        try
                        {
                            nameListForCopiedShapes.Add(shape.Name);
                        }
                        catch
                        {
                            //handling corrupted shapes
                            shape.Copy();
                            var fixedShape = _previousSlideForCopyEvent.Shapes.Paste()[1];
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
                    var range = currentSlide.Shapes.Range(nameListForPastedShapes.ToArray());
                    foreach (PowerPoint.Shape shape in range)
                    {
                        shape.Name = nameDictForPastedShapes[shape.Name];
                    }
                    range.Select();
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    var pastedSlides = selection.SlideRange.Cast<PowerPoint.Slide>().OrderBy(x => x.SlideIndex).ToList();

                    for (var i = 0; i < pastedSlides.Count; i++)
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
            var nameEnclosedInBrackets = new Regex(@"^\[\D+\s\d+\]$");
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
                var shapeTypeInName = new Regex(@"^[^\[]\D+\s(?=\d+$)");
                var shapeTypeForName1 = shapeTypeInName.Match(name1).ToString();
                var shapeTypeForName2 = shapeTypeInName.Match(name2).ToString();
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

                var copyFromRecorderPane =
                    GetPaneFromWindow(typeof(RecorderTaskPane), _copyFromWnd).Control as RecorderTaskPane;
                var activeRecorderPane = GetActivePane(typeof(RecorderTaskPane)).Control as RecorderTaskPane;

                if (activeRecorderPane == null ||
                    copyFromRecorderPane == null)
                {
                    return;
                }

                var slideRange = selection.SlideRange;
                var oriSlide = 0;

                foreach (var sld in slideRange)
                {
                    var oldSlide = PowerPointSlide.FromSlideFactory(_copiedSlides[oriSlide]);
                    var newSlide = PowerPointSlide.FromSlideFactory(sld as PowerPoint.Slide);

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

                    foreach (var sld in selection.SlideRange)
                    {
                        var slide = sld as PowerPoint.Slide;

                        _copiedSlides.Add(slide);
                    }

                    _copiedSlides.Sort((x, y) => (x.SlideIndex - y.SlideIndex));
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    _copiedShapes.Clear();
                    _previousSlideForCopyEvent = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                    _previousPptName = Application.ActivePresentation.Name;
                    foreach (var sh in selection.ShapeRange)
                    {
                        var shape = sh as PowerPoint.Shape;
                        _copiedShapes.Add(shape);
                    }

                    _copiedShapes.Sort((x, y) => (x.Id - y.Id));
                }
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
                MessageBox.Show(TextCollection.TabActivateErrorDescription, TextCollection.TabActivateErrorTitle);
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
                    if (Application.Version == OfficeVersion2010)
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
            if (PowerPointCurrentPresentationInfo.CurrentSlide == null) return;

            PowerPoint.Shape overlappingShape = null;
            int overlappingShapeZIndex = -1;

            var shapesInCurrentSlide = PowerPointCurrentPresentationInfo.CurrentSlide.Shapes;
            foreach (PowerPoint.Shape shape in shapesInCurrentSlide)
            {
                if (IsMouseWithinShape(shape)
                    && shape.ZOrderPosition > overlappingShapeZIndex)
                {
                    overlappingShape = shape;
                    overlappingShapeZIndex = shape.ZOrderPosition;
                }
            }
            if (overlappingShape != null)
                overlappingShape.Select();
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
                var selectedShapes = selection.ShapeRange;
                Native.SendMessage(
                    Process.GetCurrentProcess().MainWindowHandle,
                    (uint)Native.Message.WM_COMMAND,
                    new IntPtr(CommandOpenBackgroundFormat),
                    IntPtr.Zero
                    );
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
        # endregion

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon = new Ribbon1();
            return Ribbon;
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
