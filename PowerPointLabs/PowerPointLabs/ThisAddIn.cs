using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
        
        private readonly Dictionary<PowerPoint.DocumentWindow,
                                    List<CustomTaskPane>> _documentPaneMapper = new Dictionary<PowerPoint.DocumentWindow,
                                                                                               List<CustomTaskPane>>();
        private readonly Dictionary<PowerPoint.DocumentWindow,
                                    string> _documentHashcodeMapper = new Dictionary<PowerPoint.DocumentWindow,
                                                                                     string>();

        public Ribbon1 Ribbon;

        private const string VersionNotCompatibleMsg =
            "This file is not fully compatible with some features of PowerPointLabs because it is " +
            "in the outdated .ppt format used by PowerPoint 2007 (and older). If you wish to use the " +
            "full power of PowerPointLabs to enhance this file, please save in the .pptx format used " +
            "by PowerPoint 2010 and newer.";
        private bool _oldVersion;

        # region Powerpoint Application Event Handlers
        private void ThisAddIn_Startup(object sender, EventArgs e)
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
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += ThisAddIn_NewPresentation;
            Application.AfterNewPresentation += ThisAddIn_AfterNewPresentation;
            Application.PresentationOpen += ThisAddIn_PrensentationOpen;
            Application.PresentationClose += ThisAddIn_PresentationClose;

            // Priority Mid: Window Actions
            Application.WindowActivate += ThisAddIn_ApplicationOnWindowActivate;
            Application.WindowDeactivate += ThisAddIn_ApplicationOnWindowDeactivate;
            Application.WindowSelectionChange += ThisAddIn_SelectionChanged;
            Application.SlideShowBegin += SlideShowBeginHandler;
            Application.SlideShowEnd += SlideShowEndHandler;

            // Priority Low: Slide Actions
            Application.SlideSelectionChanged += ThisAddIn_SlideSelectionChanged;
        }

        private void ThisAddIn_ApplicationOnWindowDeactivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
        }

        private void ThisAddIn_ApplicationOnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            if (pres != null)
            {
                Ribbon._embedAudioVisible = !pres.Name.EndsWith(".ppt");
            }
        }

        private void ThisAddIn_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            Ribbon.removeCaptionsEnabled = SlidesInRangeHaveCaptions(SldRange);
            Ribbon.removeAudioEnabled = SlidesInRangeHaveAudio(SldRange);
            // update recorder pane
            if (SldRange.Count > 0)
            {
                UpdateRecorderPane(SldRange.Count, SldRange[1].SlideID);
            }
            else
            {
                UpdateRecorderPane(SldRange.Count, -1);
            }

            // in case the recorder is on event
            BreakRecorderEvents();

            // ribbon function init
            Ribbon.addAutoMotionEnabled = true;
            Ribbon.reloadAutoMotionEnabled = true;
            Ribbon.reloadSpotlight = true;
            Ribbon.highlightBulletsEnabled = true;

            if (SldRange.Count != 1)
            {
                Ribbon.addAutoMotionEnabled = false;
                Ribbon.reloadAutoMotionEnabled = false;
                Ribbon.reloadSpotlight = false;
                Ribbon.highlightBulletsEnabled = false;
            }
            else
            {
                PowerPoint.Slide tmp = SldRange[1];
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
                    Ribbon.reloadAutoMotionEnabled = false;
                if (!(tmp.Name.Contains("PPTLabsSpotlight")))
                    Ribbon.reloadSpotlight = false;
            }

            Ribbon.RefreshRibbonControl("AddAnimationButton");
            Ribbon.RefreshRibbonControl("ReloadButton");
            Ribbon.RefreshRibbonControl("ReloadSpotlightButton");
            Ribbon.RefreshRibbonControl("HighlightBulletsTextButton");
            Ribbon.RefreshRibbonControl("HighlightBulletsBackgroundButton");
            Ribbon.RefreshRibbonControl("removeCaptions");
            Ribbon.RefreshRibbonControl("removeAudio");
        }

        private void ThisAddIn_SelectionChanged(PowerPoint.Selection Sel)
        {
            Ribbon.spotlightEnabled = false;
            Ribbon.inSlideEnabled = false;
            Ribbon.zoomButtonEnabled = false;
            if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape sh = Sel.ShapeRange[1];
                if (sh.Type == Office.MsoShapeType.msoAutoShape || sh.Type == Office.MsoShapeType.msoFreeform || sh.Type == Office.MsoShapeType.msoTextBox || sh.Type == Office.MsoShapeType.msoPlaceholder
                    || sh.Type == Office.MsoShapeType.msoCallout || sh.Type == Office.MsoShapeType.msoInk || sh.Type == Office.MsoShapeType.msoGroup)
                {
                    Ribbon.spotlightEnabled = true;
                }
                if ((sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle) || sh.Type == Office.MsoShapeType.msoPicture)
                {
                    Ribbon.zoomButtonEnabled = true;
                }
                if (Sel.ShapeRange.Count > 1)
                {
                    foreach (PowerPoint.Shape tempShape in Sel.ShapeRange)
                    {
                        if (sh.Type == tempShape.Type)
                        {
                            Ribbon.inSlideEnabled = true;
                            Ribbon.zoomButtonEnabled = true;
                        }
                        if (sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType != tempShape.AutoShapeType)
                        {
                            Ribbon.inSlideEnabled = false;
                            Ribbon.zoomButtonEnabled = false;
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

        private void ThisAddIn_NewPresentation(PowerPoint.Presentation Pres)
        {
            var activeWindow = Pres.Application.ActiveWindow;
            var tempName = Pres.Name.GetHashCode().ToString();

            string tempFolderPath = Path.GetTempPath() + TempFolderNamePrefix + tempName + @"\";

            if (Directory.Exists(tempFolderPath))
            {
                if (Directory.Exists(tempFolderPath))
                {
                    Directory.Delete(tempFolderPath, true);
                }

                Directory.CreateDirectory(tempFolderPath);
            }

            _documentHashcodeMapper[activeWindow] = tempName;

            // register all task panes when new document opens
            RegisterTaskPane(new RecorderTaskPane(tempName), "Record Management", activeWindow,
                             TaskPaneVisibleValueChangedEventHandler, null);
            RegisterTaskPane(new ColorPane(), "Color Panel", activeWindow, null, null);
            RegisterTaskPane(new CustomShapePane(), "Custom Shape Management", activeWindow, null, null);
        }

        // solve new un-modified unsave problem
        private void ThisAddIn_AfterNewPresentation(PowerPoint.Presentation Pres)
        {
            //Access the BuiltInDocumentProperties so that the property storage does get created.
            object o = Pres.BuiltInDocumentProperties;
            Pres.Saved = Microsoft.Office.Core.MsoTriState.msoTrue;
        }

        private void ThisAddIn_PrensentationOpen(PowerPoint.Presentation Pres)
        {
            var activeWindow = Pres.Application.ActiveWindow;
            var tempName = Pres.Name.GetHashCode().ToString();
            var tempPath = Path.GetTempPath() + TempFolderNamePrefix + tempName + @"\";

            // as long as an existing file is opened, we need to extract embedded
            // audio files and relationship XMLs to temp folder

            // if we opened a new window, register and associate panes with the window
            if (!_documentHashcodeMapper.ContainsKey(activeWindow))
            {
                // extract the media files and relationships to a folder with presentation's
                // hash code
                if (!PrepareMediaFiles(Pres, tempPath))
                {
                    _oldVersion = true;
                    return;
                }

                _oldVersion = false;

                // register all task panes when opening documents
                RegisterTaskPane(new RecorderTaskPane(tempName), "Record Management", activeWindow,
                                 TaskPaneVisibleValueChangedEventHandler, null);
                RegisterTaskPane(new ColorPane(), "Color Panel", activeWindow, null, null);
                RegisterTaskPane(new CustomShapePane(), "Custom Shape Management", activeWindow, null, null);

                _documentHashcodeMapper[activeWindow] = tempName;
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
                if (!PrepareMediaFiles(Pres, oriTempPath))
                {
                    _oldVersion = true;
                    return;
                }

                _oldVersion = false;

                var recorderPane = GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane"));
                var recorder = recorderPane.Control as RecorderTaskPane;

                recorder.SetupListsWhenOpen();
            }
        }

        private void ThisAddIn_PresentationClose(PowerPoint.Presentation Pres)
        {
            var recorderPane = GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane"));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }

            var currentWindow = recorderPane.Window as PowerPoint.DocumentWindow;

            // make sure the close event is triggered by the window that the pane belongs to
            if (currentWindow.Presentation.Name != Pres.Name)
            {
                return;
            }

            if (Pres.Saved == Office.MsoTriState.msoTrue)
            {
                // remove task pane
                var activePanes = _documentPaneMapper[Pres.Application.ActiveWindow];
                foreach (var pane in activePanes)
                {
                    CustomTaskPanes.Remove(pane);
                }

                // remove entry from mappers
                _documentPaneMapper.Remove(Pres.Application.ActiveWindow);
                _documentHashcodeMapper.Remove(Pres.Application.ActiveWindow);
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            PPMouse.StopHook();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Exiting");
            Trace.Close();
        }
        # endregion

        # region Helper Functions
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

        public CustomTaskPane GetActivePane(Type type)
        {
            return GetPaneFromWindow(type, Application.ActiveWindow);
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
                var control = pane.Control;

                if (control.GetType() == type)
                {
                    return pane;
                }
            }

            return null;
        }

        public string GetActiveWindowTempName()
        {
            return _documentHashcodeMapper[Application.ActiveWindow];
        }

        private void RegisterTaskPane(UserControl control, string title, PowerPoint.DocumentWindow wnd,
                                      EventHandler visibleChangeEventHandler,
                                      EventHandler dockPositionChangeEventHandler)
        {
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
        }

        private void TaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            var recorderPane = GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane"));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            // trigger close form event when closing hide the pane
            if (recorderPane.Visible)
            {
                recorder.RecorderPaneClosing();
            }
        }

        private bool SlidesInRangeHaveCaptions(PowerPoint.SlideRange SldRange)
        {
            foreach (PowerPoint.Slide slide in SldRange)
            {
                PowerPointSlide pptSlide = PowerPointSlide.FromSlideFactory(slide);
                if (pptSlide.HasCaptions())
                {
                    return true;
                }
            }
            return false;
        }

        private bool SlidesInRangeHaveAudio(PowerPoint.SlideRange SldRange)
        {
            foreach (PowerPoint.Slide slide in SldRange)
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
            isInSlideShow = true;
        }

        private void SlideShowEndHandler(PowerPoint.Presentation presentation)
        {
            isInSlideShow = false;
            
            var recorderPane = GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane"));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

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
                var slides = PowerPointPresentation.Slides.ToList();

                for (int i = 0; i < recorder.AudioBuffer.Count; i++)
                {
                    if (recorder.AudioBuffer[i].Count != 0)
                    {
                        foreach (var audio in recorder.AudioBuffer[i])
                        {
                            audio.Item1.EmbedOnSlide(slides[i], audio.Item2);

                            if (Globals.ThisAddIn.Ribbon.removeAudioEnabled) continue;
                            
                            Globals.ThisAddIn.Ribbon.removeAudioEnabled = true;
                            Globals.ThisAddIn.Ribbon.RefreshRibbonControl("removeAudio");
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
            var recorderPane = GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane"));
            
            // if there's no active pane associated with the current window, return
            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;
            
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
                recorder.InitializeAudioAndScript(PowerPointPresentation.CurrentSlide, null, false);

                // if the pane is shown, refresh the pane immediately
                if (recorderPane.Visible)
                {
                    recorder.UpdateLists(id);
                }
            }
        }

        private void BreakRecorderEvents()
        {
            var recorderPane = GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane"));

            if (recorderPane == null)
            {
                return;
            }

            var recorder = recorderPane.Control as RecorderTaskPane;

            // TODO:
            // Slide change event will interrupt mci device behaviour before
            // the event raised. Now we discard the record, we may want to
            // take this record by some means.
            if (recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }
        }

        private bool PrepareMediaFiles(PowerPoint.Presentation Pres, string tempPath)
        {
            try
            {
                string presName = Pres.Name;

                if (presName.EndsWith(".ppt"))
                {
                    return false;
                }

                if (!presName.Contains(".pptx"))
                {
                    presName = Pres.Name + ".pptx";
                }

                var zipName = presName.Replace(".pptx", ".zip");
                var zipFullPath = tempPath + zipName;
                var presFullName = Pres.FullName;

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
                    ErrorDialogWrapper.ShowDialog("Error when creating temp folder", string.Empty, e);
                }
                finally
                {
                    Directory.CreateDirectory(tempPath);
                }

                // this segment is added to handle "embed on other application" issue. In this
                // case, file is not saved but has embedded audio already. We need to handle
                // it specially.
                if (Pres.Path == String.Empty)
                {
                    Pres.SaveAs(tempPath + presName);
                    presFullName = tempPath + presName;
                }

                // copy the file to temp folder and rename to zip
                try
                {
                    File.Copy(presFullName, zipFullPath);
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog("Error when accessing temp folder", string.Empty, e);
                }

                // open the zip and extract media files to temp folder
                try
                {
                    ZipStorer zip = ZipStorer.Open(zipFullPath, FileAccess.Read);

                    List<ZipStorer.ZipFileEntry> dir = zip.ReadCentralDir();
                    string pattern = @"slide(\d+)\.xml";
                    Regex regex = new Regex(pattern);

                    foreach (ZipStorer.ZipFileEntry entry in dir)
                    {
                        string name = Path.GetFileName(entry.FilenameInZip);
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
                    File.Delete(zipFullPath);
                }
                catch (Exception e)
                {
                    ErrorDialogWrapper.ShowDialog("Error when extracting", "Archived files cannot be retrieved.", e);
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Error when preparing media files", "Files cannot be linked.", e);
            }

            return true;
        }

        public bool VerifyVersion()
        {
            if (_oldVersion)
            {
                MessageBox.Show(VersionNotCompatibleMsg);
                return false;
            }

            return true;
        }
        # endregion

        # region Copy paste handlers

        private PowerPoint.DocumentWindow _copyFromWnd;

        private void AfterPasteEventHandler(PowerPoint.Selection selection)
        {
            try
            {
                PowerPoint.Slide currentSlide = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                string pptName = Application.ActivePresentation.Name;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes
                    && currentSlide.SlideID != previousSlideForCopyEvent.SlideID
                    && pptName == previousPptName)
                {
                    PowerPoint.ShapeRange pastedShapes = selection.ShapeRange;
                    List<String> nameListForPastedShapes = new List<string>();
                    Dictionary<String, String> nameDictForPastedShapes = new Dictionary<string, string>();
                    List<String> nameListForCopiedShapes = new List<string>();
                    Regex namePattern = new Regex(@"^[^\[]\D+\s\d+$");
                    List<PowerPoint.Shape> corruptedShapes = new List<PowerPoint.Shape>();

                    foreach (var shape in copiedShapes)
                    {
                        try
                        {
                            if (namePattern.IsMatch(shape.Name))
                            {
                                shape.Name = "[" + shape.Name + "]";
                            }
                            nameListForCopiedShapes.Add(shape.Name);
                        }
                        catch
                        {
                            //handling corrupted shapes
                            shape.Copy();
                            var fixedShape = previousSlideForCopyEvent.Shapes.Paste()[1];
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

                    for (int i = 1; i <= pastedShapes.Count; i++)
                    {
                        PowerPoint.Shape shape = pastedShapes[i];
                        string uniqueName = Guid.NewGuid().ToString();
                        nameDictForPastedShapes[uniqueName] = nameListForCopiedShapes[i - 1];
                        shape.Name = uniqueName;
                        nameListForPastedShapes.Add(shape.Name);
                    }
                    //Re-select pasted shapes
                    var range = currentSlide.Shapes.Range(nameListForPastedShapes.ToArray());
                    foreach (var sh in range)
                    {
                        PowerPoint.Shape shape = sh as PowerPoint.Shape;
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

        private void AfterPasteRecorderEventHandler(PowerPoint.Selection selection)
        {
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                // invalid paste event triggered because of system message loss
                if (copiedSlides.Count < 1)
                {
                    return;
                }

                // if we copied from a presentation without recorder pane or pasted to a
                // presentation without recorder pane, paste event will not be entertained
                if (!_documentPaneMapper.ContainsKey(_copyFromWnd) ||
                    _documentPaneMapper[_copyFromWnd] == null ||
                    GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane")) == null)
                {
                    return;
                }

                var copyFromRecorderPane = GetPaneFromWindow(Type.GetType("PowerPointLabs.RecorderTaskPane"), _copyFromWnd).Control as RecorderTaskPane;
                var activeRecorderPane = GetActivePane(Type.GetType("PowerPointLabs.RecorderTaskPane")).Control as RecorderTaskPane;

                var slideRange = selection.SlideRange;
                var oriSlide = 0;

                foreach (var sld in slideRange)
                {
                    var oldSlide = PowerPointSlide.FromSlideFactory(copiedSlides[oriSlide]);
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
                    copiedSlides.Clear();

                    foreach (var sld in selection.SlideRange)
                    {
                        var slide = sld as PowerPoint.Slide;

                        copiedSlides.Add(slide);
                    }

                    copiedSlides.Sort((x, y) => (x.SlideIndex - y.SlideIndex));
                }
                else if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    copiedShapes.Clear();
                    previousSlideForCopyEvent = Application.ActiveWindow.View.Slide as PowerPoint.Slide;
                    previousPptName = Application.ActivePresentation.Name;
                    foreach (var sh in selection.ShapeRange)
                    {
                        var shape = sh as PowerPoint.Shape;
                        copiedShapes.Add(shape);
                    }
                    copiedShapes.Sort((PowerPoint.Shape x, PowerPoint.Shape y) => (x.Id - y.Id));
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
            TabActivate += TabActivateEventHandler;
        }

        private Native.WinEventDelegate TabActivate;

        private IntPtr eventHook = IntPtr.Zero;

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
                Native.UnhookWinEvent(eventHook);
                eventHook = IntPtr.Zero;
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

        private bool isInSlideShow = false;

        private void SetupAfterCopyPasteHandler()
        {
            PPCopy.AfterCopy += AfterCopyEventHandler;
            PPCopy.AfterPaste += AfterPasteRecorderEventHandler;
            PPCopy.AfterPaste += AfterPasteEventHandler;
        }

        private List<PowerPoint.Shape> copiedShapes = new List<PowerPoint.Shape>();
        private List<PowerPoint.Slide> copiedSlides = new List<PowerPoint.Slide>();
        private PowerPoint.Slide previousSlideForCopyEvent;
        private string previousPptName;

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
                    const string OfficeVersion2013 = "15.0";
                    const string OfficeVersion2010 = "14.0";
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
            if (!isInSlideShow)
            {
                const int CommandOpenBackgroundFormat = 0x8F;
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
                if (!isInSlideShow)
                {
                    string Shortcut_Alt_H_O = "%ho";
                    if (eventHook == IntPtr.Zero)
                    {
                        //Check whether Home tab is enabled or not
                        eventHook = Native.SetWinEventHook(
                            (uint) Native.Event.EVENT_SYSTEM_MENUEND,
                            (uint) Native.Event.EVENT_OBJECT_CREATE,
                            IntPtr.Zero,
                            TabActivate,
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

        #endregion

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
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
