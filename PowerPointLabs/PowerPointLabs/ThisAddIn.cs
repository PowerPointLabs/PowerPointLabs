using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Diagnostics;
using Microsoft.Office.Tools;
using PPExtraEventHelper;
using System.IO.Compression;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using PowerPointLabs.XMLMisc;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Deployment.Application;

namespace PowerPointLabs
{
    public partial class ThisAddIn
    {
        public Ribbon1 ribbon;
        public ArrayList indicators = new ArrayList();
        
        internal CustomTaskPane ActivateCustomTaskPane;
        internal Dictionary<PowerPoint.DocumentWindow, CustomTaskPane> documentPaneMapper = new Dictionary<PowerPoint.DocumentWindow, CustomTaskPane>();
        internal Dictionary<PowerPoint.DocumentWindow, string> documentHashcodeMapper = new Dictionary<PowerPoint.DocumentWindow, string>();
        internal bool customTaskPaneInit = false;
        internal InShowControl inShowControlBox;

        private const string TempFolderNamePrefix = @"\PowerPointLabs Temp\";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SetupLogger();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Started");
            
            PPMouse.Init(Application);
            PPCopy.Init(Application);
            SetupDoubleClickHandler();
            SetupTabActivateHandler();
            SetupAfterCopyPasteHandler();

            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += ThisAddIn_NewPresentation;
            Application.AfterNewPresentation += ThisAddIn_AfterNewPresentation;
            Application.WindowSelectionChange += ThisAddIn_SelectionChanged;
            Application.PresentationClose += ThisAddIn_PresentationClose;
            Application.PresentationOpen += ThisAddIn_PrensentationOpen;
            Application.WindowActivate += ThisAddIn_ApplicationOnWindowActivate;
            Application.WindowDeactivate += ThisAddIn_ApplicationOnWindowDeactivate;
            Application.SlideSelectionChanged += ThisAddIn_SlideSelectionChanged;

            Application.SlideShowBegin += SlideShowBeginHandler;
            Application.SlideShowEnd += SlideShowEndHandler;
        }

        private void ThisAddIn_ApplicationOnWindowDeactivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
        }

        private void ThisAddIn_ApplicationOnWindowActivate(PowerPoint.Presentation pres, PowerPoint.DocumentWindow wn)
        {
            // set recorder pane for current window as active recorder pane
            if (documentPaneMapper.ContainsKey(wn))
            {
                ActivateCustomTaskPane = documentPaneMapper[wn];
            }
        }

        void SetupLogger()
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

        void SetupRecorderTaskPane(PowerPoint.DocumentWindow wnd)
        {
            try
            {
                var tempName = wnd.Presentation.Name.GetHashCode().ToString();
                var recorderPane = new RecorderTaskPane(tempName);
                var width = recorderPane.Width;

                // register the recorder task pane to the CustomTaskPanes collection
                ActivateCustomTaskPane = CustomTaskPanes.Add(recorderPane, "Record Script", wnd);

                // map the current window with the task pane
                documentPaneMapper[wnd] = ActivateCustomTaskPane;
                documentHashcodeMapper[wnd] = tempName;

                // recorder task pane customization
                // currently recorder pane is always visible since only one pane in the
                // custom task pane collection
                ActivateCustomTaskPane.Visible = false;
                ActivateCustomTaskPane.VisibleChanged += TaskPaneVisibleValueChangedEventHandler;
                ActivateCustomTaskPane.Width = width + 20;
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Error when setting up panes",
                                              "Some working panes cannot be loaded successfully.", e);
                throw;
            }
        }

        public string GetActiveWindowTempName()
        {
            return documentHashcodeMapper[Application.ActiveWindow];
        }

        void TaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            var recorder = ActivateCustomTaskPane.Control as RecorderTaskPane;
            // hide the pane
            
            if (ActivateCustomTaskPane.Visible)
            {
                recorder.RecorderPaneClosing();
            }
        }

        void ThisAddIn_SelectionChanged(PowerPoint.Selection Sel)
        {
            ribbon.spotlightEnabled = false;
            ribbon.inSlideEnabled = false;
            ribbon.zoomButtonEnabled = false;
            if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape sh = Sel.ShapeRange[1];
                if (sh.Type == Office.MsoShapeType.msoAutoShape || sh.Type == Office.MsoShapeType.msoFreeform || sh.Type == Office.MsoShapeType.msoTextBox || sh.Type == Office.MsoShapeType.msoPlaceholder
                    || sh.Type == Office.MsoShapeType.msoCallout || sh.Type == Office.MsoShapeType.msoInk || sh.Type == Office.MsoShapeType.msoGroup)
                {
                    ribbon.spotlightEnabled = true;
                }
                if ((sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle) || sh.Type == Office.MsoShapeType.msoPicture)
                {
                    ribbon.zoomButtonEnabled = true;
                }
                if (Sel.ShapeRange.Count > 1)
                {
                    foreach (PowerPoint.Shape tempShape in Sel.ShapeRange)
                    {
                        if (sh.Type == tempShape.Type)
                        {
                            ribbon.inSlideEnabled = true;
                            ribbon.zoomButtonEnabled = true;
                        }
                        if (sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType != tempShape.AutoShapeType)
                        {
                            ribbon.inSlideEnabled = false;
                            ribbon.zoomButtonEnabled = false;
                            break;
                        }
                    }
                }
            }

            ribbon.RefreshRibbonControl("AddSpotlightButton");
            ribbon.RefreshRibbonControl("InSlideAnimateButton");
            ribbon.RefreshRibbonControl("AddZoomInButton");
            ribbon.RefreshRibbonControl("AddZoomOutButton");
            ribbon.RefreshRibbonControl("ZoomToAreaButton");
        }

        bool SlidesInRangeHaveCaptions(PowerPoint.SlideRange SldRange)
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

        bool SlidesInRangeHaveAudio(PowerPoint.SlideRange SldRange)
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

        void ThisAddIn_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            ribbon.removeCaptionsEnabled = SlidesInRangeHaveCaptions(SldRange);
            ribbon.removeAudioEnabled = SlidesInRangeHaveAudio(SldRange);
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
            ribbon.addAutoMotionEnabled = true;
            ribbon.reloadAutoMotionEnabled = true;
            ribbon.reloadSpotlight = true;
            ribbon.highlightBulletsEnabled = true;

            if (SldRange.Count != 1)
            {
                ribbon.addAutoMotionEnabled = false;
                ribbon.reloadAutoMotionEnabled = false;
                ribbon.reloadSpotlight = false;
                ribbon.highlightBulletsEnabled = false;
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
                    ribbon.reloadAutoMotionEnabled = false;
                if (!(tmp.Name.Contains("PPTLabsSpotlight")))
                    ribbon.reloadSpotlight = false;
            }

            ribbon.RefreshRibbonControl("AddAnimationButton");
            ribbon.RefreshRibbonControl("ReloadButton");
            ribbon.RefreshRibbonControl("ReloadSpotlightButton");
            ribbon.RefreshRibbonControl("HighlightBulletsTextButton");
            ribbon.RefreshRibbonControl("HighlightBulletsBackgroundButton");
            ribbon.RefreshRibbonControl("removeCaptions");
            ribbon.RefreshRibbonControl("removeAudio");
        }

        // TODO:
        // Each new presentation should be assigned a UID, this will be used
        // to distinguish different audio files with same slide number but
        // within different presentation.
        void ThisAddIn_NewPresentation(PowerPoint.Presentation Pres)
        {
            try
            {
                string tempFolderPath = Path.GetTempPath() + TempFolderNamePrefix + Pres.Name.GetHashCode().ToString() + @"\";

                if (Directory.Exists(tempFolderPath))
                {
                    Directory.Delete(tempFolderPath, true);
                }

                Directory.CreateDirectory(tempFolderPath);

                // setup a new recorder pane when an exist file opened
                SetupRecorderTaskPane(Pres.Application.ActiveWindow);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Error when opening new file",
                                              "New presentation cannot be loaded successfully.", e);
                throw;
            }
        }

        // solve new un-modified unsave problem
        void ThisAddIn_AfterNewPresentation(PowerPoint.Presentation Pres)
        {
            //Access the BuiltInDocumentProperties so that the property storage does get created.
            object o = Pres.BuiltInDocumentProperties;
            Pres.Saved = Microsoft.Office.Core.MsoTriState.msoTrue;
        }

        void ThisAddIn_PrensentationOpen(PowerPoint.Presentation Pres)
        {
            try
            {
                // extract embedded audio files to temp folder
                PrepareMediaFiles(Pres);
                // setup a new recorder pane when an exist file opened
                SetupRecorderTaskPane(Pres.Application.ActiveWindow);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Error opening file.", "File cannot be opened.", e);
                throw;
            }
        }

        void ThisAddIn_PresentationClose(PowerPoint.Presentation Pres)
        {
            var recorder = ActivateCustomTaskPane.Control as RecorderTaskPane;
            
            if (recorder.HasEvent())
            {
                recorder.ForceStopEvent();
            }
            
            var currentWindow = ActivateCustomTaskPane.Window as PowerPoint.DocumentWindow;

            // make sure the close event is triggered by the window that the pane belongs to
            if (currentWindow.Presentation.Name != Pres.Name)
            {
                return;
            }

            if (Pres.Saved == Office.MsoTriState.msoTrue)
            {
                // remove task pane
                CustomTaskPanes.Remove(documentPaneMapper[Pres.Application.ActiveWindow]);

                // remove entry from mappers
                documentPaneMapper.Remove(Pres.Application.ActiveWindow);
                documentHashcodeMapper.Remove(Pres.Application.ActiveWindow);
            }
        }

        private void SlideShowBeginHandler(PowerPoint.SlideShowWindow wn)
        {
            isInSlideShow = true;
        }

        private void SlideShowEndHandler(PowerPoint.Presentation presentation)
        {
            isInSlideShow = false;
            var recorderPane = ActivateCustomTaskPane.Control as RecorderTaskPane;

            // force recording session ends
            if (recorderPane.HasEvent())
            {
                recorderPane.ForceStopEvent();
            }

            // enable slide show button
            recorderPane.EnableSlideShow();

            // when leave the show, dispose the in-show control if we have one
            recorderPane.DisposeInSlideControlBox();

            // if audio buffer is not empty, render the effects
            if (recorderPane.AudioBuffer.Count != 0)
            {
                var slides = PowerPointPresentation.Slides.ToList();

                for (int i = 0; i < recorderPane.AudioBuffer.Count; i++)
                {
                    if (recorderPane.AudioBuffer[i].Count != 0)
                    {
                        foreach (var audio in recorderPane.AudioBuffer[i])
                        {
                            audio.Item1.EmbedOnSlide(slides[i], audio.Item2);

                            if (Globals.ThisAddIn.ribbon.removeAudioEnabled) continue;
                            
                            Globals.ThisAddIn.ribbon.removeAudioEnabled = true;
                            Globals.ThisAddIn.ribbon.RefreshRibbonControl("removeAudio");
                        }
                    }
                }
            }

            // clear the buffer after embed
            recorderPane.AudioBuffer.Clear();

            // change back the slide range settings
            Application.ActivePresentation.SlideShowSettings.RangeType = PowerPoint.PpSlideShowRangeType.ppShowAll;
        }

        private void UpdateRecorderPane(int count, int id)
        {
            var recorderPane = ActivateCustomTaskPane.Control as RecorderTaskPane;
            // if the user has selected none or more than 1 slides, recorder pane should show nothing
            if (count != 1)
            {
                if (ActivateCustomTaskPane.Visible)
                {
                    recorderPane.ClearDisplayLists();
                }
            }
            else
            {
                // initailize the current slide
                recorderPane.InitializeAudioAndScript(PowerPointPresentation.CurrentSlide, null, false);

                // if the pane is shown, refresh the pane immediately
                if (ActivateCustomTaskPane.Visible)
                {
                    recorderPane.UpdateLists(id);
                }
            }
        }

        private void BreakRecorderEvents()
        {
            var recorderPane = ActivateCustomTaskPane.Control as RecorderTaskPane;

            // TODO:
            // Slide change event will interrupt mci device behaviour before
            // the event raised. Now we discard the record, we may want to
            // take this record by some means.
            if (recorderPane.HasEvent())
            {
                recorderPane.ForceStopEvent();
            }
        }

        private void PrepareMediaFiles(PowerPoint.Presentation Pres)
        {
            try
            {
                string presName = Pres.Name;

                if (!presName.Contains(".pptx"))
                {
                    presName = Pres.Name + ".pptx";
                }

                string tempPath = Path.GetTempPath() + TempFolderNamePrefix + Pres.Name.GetHashCode().ToString() + @"\";
                string zipName = presName.Replace(".pptx", ".zip");
                string zipFullPath = tempPath + zipName;
                string presFullName = Pres.FullName;

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
                    ErrorDialogWrapper.ShowDialog("Error when creating temp folder", "Temp folder " + tempPath + " cannot be created.", e);
                    throw;
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
                    throw;
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
                    throw;
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Error when preparing media files", "Files cannot be linked.", e);
                throw;
            }
        }

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

                var copyFromRecorderPane = documentPaneMapper[_copyFromWnd].Control as RecorderTaskPane;
                var activeRecorderPane = ActivateCustomTaskPane.Control as RecorderTaskPane;

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

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            PPExtraEventHelper.PPMouse.StopHook();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Exiting");
            Trace.Close();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon1();
            return ribbon;
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
