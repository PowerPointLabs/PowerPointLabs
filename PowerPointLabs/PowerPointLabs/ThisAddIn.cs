using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Tools;
using PPExtraEventHelper;
using System.IO.Compression;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointLabs
{
    public partial class ThisAddIn
    {
        private const string TempFolderNamePrefix = @"\PowerPointLabs Temp\";

        private readonly string _defaultShapeMasterFolderPrefix =
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private const string DefaultShapeMasterFolderName = @"\PowerPointLabs Custom Shapes";
        private const string DefaultShapeCategoryName = @"My Shapes";
        private const string ShapeGalleryPptxName = @"ShapeGallery";
        
        private bool _oldVersion;

        private readonly Dictionary<PowerPoint.DocumentWindow,
                                    List<CustomTaskPane>> _documentPaneMapper = new Dictionary<PowerPoint.DocumentWindow,
                                                                                               List<CustomTaskPane>>();
        private readonly Dictionary<PowerPoint.DocumentWindow,
                                    string> _documentHashcodeMapper = new Dictionary<PowerPoint.DocumentWindow,
                                                                                     string>();

        internal PowerPointShapeGalleryPresentation ShapePresentation;
        
        public Ribbon1 Ribbon;

        # region Powerpoint Application Event Handlers
        private void ThisAddInStartup(object sender, EventArgs e)
        {
            SetupLogger();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Started");
            
            PPMouse.Init(Application);
            PPCopy.Init(Application);
            SetupDoubleClickHandler();
            SetupTabActivateHandler();
            SetupAfterCopyPasteHandler();

            // According to MSDN, when more than 1 event are triggered, callback's invoking sequence
            // follows the defining order. I.e. the earlier you defined, the earlier it will be
            // executed.

            // Here, we want the priority to be: Application action > Window action > Slide action

            // Priority High: Application Actions
            ((PowerPoint.EApplication_Event)Application).NewPresentation += ThisAddInNewPresentation;
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
            // in this case, we are closing the last client presentation, therefore we
            // we can close the shape gallery
            if (Application.Presentations.Count == 2 &&
                ShapePresentation.Opened)
            {
                ShapePresentation.Close();
            }
        }

        private void ThisAddInApplicationOnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            if (pres != null)
            {
                Ribbon.EmbedAudioVisible = !pres.Name.EndsWith(".ppt");
            }
        }

        private void ThisAddInSlideSelectionChanged(PowerPoint.SlideRange sldRange)
        {
            Ribbon.RemoveCaptionsEnabled = SlidesInRangeHaveCaptions(sldRange);
            Ribbon.RemoveAudioEnabled = SlidesInRangeHaveAudio(sldRange);
            // update recorder pane
            if (sldRange.Count > 0)
            {
                UpdateRecorderPane(sldRange.Count, sldRange[1].SlideID);
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
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
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
                if (sh.Type == Office.MsoShapeType.msoAutoShape || sh.Type == Office.MsoShapeType.msoFreeform || sh.Type == Office.MsoShapeType.msoTextBox || sh.Type == Office.MsoShapeType.msoPlaceholder
                    || sh.Type == Office.MsoShapeType.msoCallout || sh.Type == Office.MsoShapeType.msoInk || sh.Type == Office.MsoShapeType.msoGroup)
                {
                    Ribbon.SpotlightEnabled = true;
                }
                if ((sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle) || sh.Type == Office.MsoShapeType.msoPicture)
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
            }

            Ribbon.RefreshRibbonControl("AddSpotlightButton");
            Ribbon.RefreshRibbonControl("InSlideAnimateButton");
            Ribbon.RefreshRibbonControl("AddZoomInButton");
            Ribbon.RefreshRibbonControl("AddZoomOutButton");
            Ribbon.RefreshRibbonControl("ZoomToAreaButton");
        }

        private void ThisAddInNewPresentation(PowerPoint.Presentation pres)
        {
            var tempName = pres.Name.GetHashCode().ToString(CultureInfo.InvariantCulture);
            var tempFolderPath = Path.GetTempPath() + TempFolderNamePrefix + tempName + @"\";

            PrepareTempFolder(tempFolderPath);
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
            var tempPath = Path.GetTempPath() + TempFolderNamePrefix + tempName + @"\";

            // as long as an existing file is opened, we need to extract embedded
            // audio files and relationship XMLs to temp folder

            // if we opened a new window, register and associate panes with the window
            if (!_documentHashcodeMapper.ContainsKey(activeWindow))
            {
                // extract the media files and relationships to a folder with presentation's
                // hash code
                if (!PrepareMediaFiles(pres, tempPath))
                {
                    _oldVersion = true;
                    return;
                }

                _oldVersion = false;
            }
            else
            {
                // this case happens when we create a new blank presentation, and open
                // an exisiting file immediately. The exsiting file shares the same
                // window with the blank presentation, but the blank presentation has
                // gone without triggering ApplicationClose event. Instead,
                // ApplicationOpen and SlideSlectionChange events are triggered.

                // to deal with this special case, we need to prepare media files and
                // xml relationships to the folder belongs to the blank presentation, and
                // manually call the setup method of the recorder pane.
                var oriTempPath = Path.GetTempPath() + TempFolderNamePrefix +
                                  _documentHashcodeMapper[activeWindow] + @"\";
                
                if (!PrepareMediaFiles(pres, oriTempPath))
                {
                    _oldVersion = true;
                    return;
                }

                _oldVersion = false;

                var recorderPane = GetActivePane(typeof(RecorderTaskPane));

                if (recorderPane == null)
                {
                    return;
                }

                var recorder = recorderPane.Control as RecorderTaskPane;

                if (recorder == null)
                {
                    return;
                }

                recorder.SetupListsWhenOpen();
            }
        }

        private void ThisAddInPresentationClose(PowerPoint.Presentation pres)
        {
            var recorderPane = GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder != null &&
                recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }

            var currentWindow = recorderPane.Window as PowerPoint.DocumentWindow;

            // make sure the close event is triggered by the window that the pane belongs to
            if (currentWindow != null &&
                currentWindow.Presentation.Name != pres.Name)
            {
                return;
            }

            if (pres.Saved == Office.MsoTriState.msoTrue)
            {
                // remove task pane
                var activePanes = _documentPaneMapper[pres.Application.ActiveWindow];
                foreach (var pane in activePanes)
                {
                    CustomTaskPanes.Remove(pane);
                }

                // remove entry from mappers
                _documentPaneMapper.Remove(pres.Application.ActiveWindow);
                _documentHashcodeMapper.Remove(pres.Application.ActiveWindow);
            }
        }

        private void ThisAddInShutdown(object sender, EventArgs e)
        {
            PPMouse.StopHook();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Exiting");
            Trace.Close();
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
            return GetPaneFromWindow(type, Application.ActiveWindow);
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
            if (ShapePresentation != null) return;

            ShapePresentation =
                new PowerPointShapeGalleryPresentation(_defaultShapeMasterFolderPrefix + DefaultShapeMasterFolderName,
                                                       ShapeGalleryPptxName);

            ShapePresentation.Open(withWindow: false, focus: false);
            ShapePresentation.AddCategory(DefaultShapeCategoryName);
            ShapePresentation.Save();
        }

        public void RegisterRecorderPane(PowerPoint.Presentation presentation)
        {
            if (GetActivePane(typeof(RecorderTaskPane)) != null)
            {
                return;
            }

            var activeWindow = presentation.Application.ActiveWindow;
            var tempName = presentation.Name.GetHashCode().ToString(CultureInfo.InvariantCulture);

            _documentHashcodeMapper[activeWindow] = tempName;

            RegisterTaskPane(new RecorderTaskPane(tempName), "Record Management", activeWindow,
                             TaskPaneVisibleValueChangedEventHandler, null);
        }

        public void RegisterColorPane(PowerPoint.Presentation presentation)
        {
            if (GetActivePane(typeof(ColorPane)) != null)
            {
                return;
            }

            var activeWindow = presentation.Application.ActiveWindow;

            TaskPaneSetup(presentation);
            RegisterTaskPane(new ColorPane(), "Color Panel", activeWindow, null, null);
        }

        public void RegisterCustomShapePane(PowerPoint.Presentation presentation)
        {
            if (GetActivePane(typeof(CustomShapePane)) != null)
            {
                return;
            }

            var activeWindow = presentation.Application.ActiveWindow;
            
            TaskPaneSetup(presentation);
            RegisterTaskPane(
                new CustomShapePane(_defaultShapeMasterFolderPrefix + DefaultShapeMasterFolderName,
                                    DefaultShapeCategoryName), "Shapes Lab", activeWindow, null, null);
        }

        public void SyncShapeAdd(string shapeName, string shapeFullName)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow) continue;

                var shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl == null) continue;

                shapePaneControl.AddCustomShape(shapeName, shapeFullName, false);
            }
        }

        public void SyncShapeRemove(string shapeName)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow) continue;

                var shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl == null) continue;

                shapePaneControl.RemoveCustomShape(shapeName);
            }
        }

        public void SyncShapeRename(string shapeOldName, string shapeNewName)
        {
            foreach (PowerPoint.DocumentWindow window in Globals.ThisAddIn.Application.Windows)
            {
                if (window == Application.ActiveWindow) continue;

                var shapePaneControl = GetControlFromWindow(typeof(CustomShapePane), window) as CustomShapePane;

                if (shapePaneControl == null) continue;

                shapePaneControl.RenameCustomShape(shapeOldName, shapeNewName);
            }
        }
        # endregion

        # region Helper Functions
        private const string SlideXmlSearchPattern = @"slide(\d+)\.xml";
        
        private void SetupLogger()
        {
            // The folder for the roaming current user 
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            // Combine the base folder with your specific folder....
            string specificFolder = Path.Combine(folder, "PowerPointLabs");

            // Check if folder exists and if not, create it
            if (!Directory.Exists(specificFolder))
                Directory.CreateDirectory(specificFolder);
            string fileName = Path.Combine(specificFolder, "PowerPointLabs_Log_1.log");

            Trace.AutoFlush = true;
            Trace.Listeners.Add(new TextWriterTraceListener(fileName));
        }

        private void RegisterTaskPane(UserControl control, string title, PowerPoint.DocumentWindow wnd,
                                      EventHandler visibleChangeEventHandler,
                                      EventHandler dockPositionChangeEventHandler)
        {
            var loadingDialog = new LoadingDialog();
            loadingDialog.Show();
            loadingDialog.Refresh();

            // note down the control's width
            var width = control.Width;

            // register the user control to the CustomTaskPanes collection and set it as
            // current active task pane;
            var taskPane = CustomTaskPanes.Add(control, title, wnd);

            // map the current window with the task pane
            if (!_documentPaneMapper.ContainsKey(wnd))
            {
                _documentPaneMapper[wnd] = new List<CustomTaskPane>();
            }

            _documentPaneMapper[wnd].Add(taskPane);

            // task pane UI setup
            taskPane.Visible = false;
            taskPane.Width = width + 20;

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
        }

        private void TaskPaneSetup(PowerPoint.Presentation presentation)
        {
            var activeWindow = presentation.Application.ActiveWindow;
            var tempName = presentation.Name.GetHashCode().ToString(CultureInfo.InvariantCulture);

            _documentHashcodeMapper[activeWindow] = tempName;
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
            if (recorder != null && recorderPane.Visible)
            {
                recorder.RecorderPaneClosing();
            }
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
        }

        private void SlideShowEndHandler(PowerPoint.Presentation presentation)
        {
            _isInSlideShow = false;
            
            var recorderPane = GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null)
            {
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
                var slides = PowerPointCurrentPresentationInfo.Slides.ToList();

                for (int i = 0; i < recorder.AudioBuffer.Count; i++)
                {
                    if (recorder.AudioBuffer[i].Count != 0)
                    {
                        foreach (var audio in recorder.AudioBuffer[i])
                        {
                            audio.Item1.EmbedOnSlide(slides[i], audio.Item2);

                            if (Globals.ThisAddIn.Ribbon.RemoveAudioEnabled) continue;
                            
                            Globals.ThisAddIn.Ribbon.RemoveAudioEnabled = true;
                            Globals.ThisAddIn.Ribbon.RefreshRibbonControl("RemoveAudioButton");
                        }
                    }
                }
            }

            // clear the buffer after embed
            recorder.AudioBuffer.Clear();

            // change back the slide range settings
            Application.ActivePresentation.SlideShowSettings.RangeType = PowerPoint.PpSlideShowRangeType.ppShowAll;
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

        private void BreakRecorderEvents()
        {
            var recorderPane = GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            // TODO:
            // Slide change event will interrupt mci device behaviour before
            // the event raised. Now we discard the record, we may want to
            // take this record by some means.
            if (recorder != null &&
                recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }
        }

        private bool PrepareMediaFiles(PowerPoint.Presentation pres, string tempPath)
        {
            try
            {
                string presName = pres.Name;

                if (presName.EndsWith(".ppt"))
                {
                    return false;
                }

                if (!presName.Contains(".pptx"))
                {
                    presName = pres.Name + ".pptx";
                }

                var zipName = presName.Replace(".pptx", ".zip");
                var zipFullPath = tempPath + zipName;
                var presFullName = pres.FullName;

                // before we do everything, check if there's an undelete old zip file
                // due to some error
                if (File.Exists(zipFullPath))
                {
                    File.SetAttributes(zipFullPath, FileAttributes.Normal);
                    File.Delete(zipFullPath);
                }

                // if temp folder exists, delete then create in case 2 different files
                // share the same name
                PrepareTempFolder(tempPath);

                // this segment is added to handle "embed on other application" issue. In this
                // case, file is not saved but has embedded audio already. We need to handle
                // it specially.
                if (pres.Path == String.Empty)
                {
                    pres.SaveAs(tempPath + presName);
                    presFullName = tempPath + presName;
                }

                // copy the file to temp folder and rename to zip
                try
                {
                    File.Copy(presFullName, zipFullPath);
                    File.SetAttributes(zipFullPath, FileAttributes.Normal);
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog(TextCollection.AccessTempFolderErrorMsg, string.Empty, e);
                }

                var fileInfo = new FileInfo(zipFullPath);
                
                // if there's nothing inside the zip file, we do nothing
                if (fileInfo.Length == 0)
                {
                    return true;
                }

                // open the zip and extract media files to temp folder
                try
                {
                    var zip = ZipStorer.Open(zipFullPath, FileAccess.Read);
                    var dir = zip.ReadCentralDir();

                    var regex = new Regex(SlideXmlSearchPattern);

                    foreach (var entry in dir)
                    {
                        var name = Path.GetFileName(entry.FilenameInZip);

                        if (name == null) continue;

                        if (name.Contains(".wav"))
                        {
                            zip.ExtractFile(entry, tempPath + name);
                        }
                        else if (regex.IsMatch(name))
                        {
                            zip.ExtractFile(entry, tempPath + name);

                            //var match = regex.Match(name);
                        }
                    }

                    zip.Close();
                    File.SetAttributes(zipFullPath, FileAttributes.Normal);
                    File.Delete(zipFullPath);
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog(TextCollection.ExtraErrorMsg, "Archived files cannot be retrieved.", e);
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog(TextCollection.PrepareMediaErrorMsg, "Files cannot be linked.", e);
            }

            return true;
        }

        private void PrepareTempFolder(string tempPath)
        {
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
        }

        public bool VerifyVersion()
        {
            if (_oldVersion)
            {
                MessageBox.Show(TextCollection.VersionNotCompatibleMsg);
                return false;
            }

            return true;
        }
        # endregion

        # region Copy paste handlers

        private PowerPoint.DocumentWindow _copyFromWnd;
        private Regex _shapeNamePattern = new Regex(@"^[^\[]\D+\s\d+$");
        private HashSet<String> isShapeMatchedAlready; 

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
                    var namePattern = new Regex(@"^[^\[]\D+\s\d+$");
                    var corruptedShapes = new List<PowerPoint.Shape>();

                    foreach (var shape in _copiedShapes)
                    {
                        try
                        {
                            if (_shapeNamePattern.IsMatch(shape.Name))
                            {
                                shape.Name = "[" + shape.Name + "]";
                            }
                            nameListForCopiedShapes.Add(shape.Name);
                        }
                        catch
                        {
                            //handling corrupted shapes
                            shape.Copy();
                            var fixedShape = _previousSlideForCopyEvent.Shapes.Paste()[1];
                            fixedShape.Name = "[" + shape.Name + "]";
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

                    for (int i = 0; i < corruptedShapes.Count; i++)
                    {
                        corruptedShapes[i].Delete();
                    }

                    isShapeMatchedAlready = new HashSet<string>();
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
                    && shape.Left == _copiedShapes[i].Left
                    && shape.Height == _copiedShapes[i].Height
                    && !isShapeMatchedAlready.Contains(_copiedShapes[i].Id.ToString()))
                {
                    isShapeMatchedAlready.Add(_copiedShapes[i].Id.ToString());
                    return i;
                }
            }
            //Blur matching:
            for (int i = 0; i < _copiedShapes.Count; i++)
            {
                if (IsSimilarShape(shape, _copiedShapes[i])
                    && IsSimilarName(shape.Name, _copiedShapes[i].Name)
                    && !isShapeMatchedAlready.Contains(_copiedShapes[i].Id.ToString()))
                {
                    isShapeMatchedAlready.Add(_copiedShapes[i].Id.ToString());
                    return i;
                }
            }
            return -1;
        }

        private bool IsSimilarShape(PowerPoint.Shape shape, PowerPoint.Shape shape2)
        {
            return shape.Width == shape2.Width
                   && shape.Height == shape2.Height
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
                    GetPaneFromWindow(typeof (RecorderTaskPane), _copyFromWnd).Control as RecorderTaskPane;
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
                string description = "To activate 'Double Click to Open Property' feature, you need to enable 'Home' tab " +
                              "in Options -> Customize Ribbon -> Main Tabs -> tick the checkbox of 'Home' -> click OK but" +
                              "ton to save.";
                string title = "Unable to activate 'Double Click to Open Property' feature";
                MessageBox.Show(description, title);
            }
        }

        #endregion

        #region Double Click to Open Property Window

        private const string OfficeVersion2013 = "15.0";
        private const string OfficeVersion2010 = "14.0";
        
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
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    if (Application.Version == OfficeVersion2013)
                    {
                        OpenPropertyWindowForOffice13(selection);
                    }
                    else if (Application.Version == OfficeVersion2010)
                    {
                        OpenPropertyWindowForOffice10();
                    }
                }
            }
            catch (COMException e)
            {
                string logText = "DoubleClickEventHandler" + ": " + e.Message + ": " + e.StackTrace;
                Trace.TraceError(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            }
        }

        //For office 2013 only:
        //Open Background Format window, then selecting the shape will
        //convert the window to Property window
        private void OpenPropertyWindowForOffice13(PowerPoint.Selection selection)
        {
            if (!_isInSlideShow)
            {
                var selectedShapes = selection.ShapeRange;
                Native.SendMessage(
                    Process.GetCurrentProcess().MainWindowHandle,
                    (uint) Native.Message.WM_COMMAND,
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
                    string Shortcut_Alt_H_O = "%ho";
                    if (_eventHook == IntPtr.Zero)
                    {
                        //Check whether Home tab is enabled or not
                        _eventHook = Native.SetWinEventHook(
                            (uint) Native.Event.EVENT_SYSTEM_MENUEND,
                            (uint) Native.Event.EVENT_OBJECT_CREATE,
                            IntPtr.Zero,
                            _tabActivate,
                            (uint) Process.GetCurrentProcess().Id,
                            0,
                            0);
                    }
                    SendKeys.Send(Shortcut_Alt_H_O);
                }
            }
            catch (InvalidOperationException)
            {
                //
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
