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
        
        internal Microsoft.Office.Tools.CustomTaskPane customTaskPane;
        internal RecorderTaskPane recorderTaskPane;
        internal InShowControl inShowControlBox;

        private const int _taskPaneWidth = 300;
        private const string TempFolderName = @"\PowerPointLabs Temp\";
        private const string SpeechShapePrefix = "PowerPointLabs Speech";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SetupLogger();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Started");
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += new Microsoft.Office.Interop.PowerPoint.EApplication_NewPresentationEventHandler(ThisAddIn_NewPresentation);
            ((PowerPoint.EApplication_Event)this.Application).WindowSelectionChange += new Microsoft.Office.Interop.PowerPoint.EApplication_WindowSelectionChangeEventHandler(ThisAddIn_SelectionChanged);
            ((PowerPoint.EApplication_Event)this.Application).SlideSelectionChanged += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideSelectionChangedEventHandler(ThisAddIn_SlideSelectionChanged);
            ((PowerPoint.EApplication_Event)this.Application).PresentationClose += new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationCloseEventHandler(ThisAddIn_PresentationClose);
            ((PowerPoint.EApplication_Event)this.Application).PresentationOpen += new Microsoft.Office.Interop.PowerPoint.EApplication_PresentationOpenEventHandler(ThisAddIn_PrensentationOpen);

            Application.SlideShowBegin += SlideShowBeginHandler;
            Application.SlideShowEnd += SlideShowEndHandler;
            PPMouse.Init(Application);
            PPCopy.Init(Application);
            SetupDoubleClickHandler();
            SetupTabActivateHandler();
            SetupAfterCopyPasteHandler();
            SetupRecorderTaskPane();
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

        void SetupRecorderTaskPane()
        {
            // register the recorder task pane to the CustomTaskPanes collection
            recorderTaskPane = new RecorderTaskPane();
            customTaskPane = CustomTaskPanes.Add(recorderTaskPane, "Record Script");

            // recorder task pane customization
            // currently recorder pane is always visible since only one pane in the
            // custom task pane collection
            customTaskPane.Visible = false;
            customTaskPane.VisibleChanged += TaskPaneVisibleValueChangedEventHandler;
            customTaskPane.Width = _taskPaneWidth;
        }

        void TaskPaneVisibleValueChangedEventHandler(object sender, EventArgs e)
        {
            // hide the pane
            if (ribbon._recorderPaneVisible)
            {
                ribbon._recorderPaneVisible = false;
                recorderTaskPane.RecorderPaneClosing();
            }
            else
            {
                ribbon._recorderPaneVisible = true;
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

        void ThisAddIn_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            ribbon.removeCaptionsEnabled = SlidesInRangeHaveCaptions(SldRange);

            // update recorder pane
            UpdateRecorderPane(SldRange.Count, SldRange[1].SlideID);

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
        }

        // TODO:
        // Each new presentation should be assigned a UID, this will be used
        // to distinguish different audio files with same slide number but
        // within different presentation.
        void ThisAddIn_NewPresentation(PowerPoint.Presentation Pres)
        {
        }

        void ThisAddIn_PrensentationOpen(PowerPoint.Presentation Pres)
        {
            // before open, check if the temp folder already exists. If it is, delete it
            string tempFolderPath = Path.GetTempPath() + TempFolderName;
            if (Directory.Exists(tempFolderPath))
            {
                Directory.Delete(tempFolderPath, true);
            }

            // extract embedded audio files to temp folder
            PrepareMediaFiles(Pres);
            // set up recorder pane
            recorderTaskPane.SetupListsWhenOpen();
        }

        void ThisAddIn_PresentationClose(PowerPoint.Presentation Pres)
        {
            var saved = Pres.Saved;

            // embed all audios once again to preserve the playing sequence
            try
            {
                recorderTaskPane.ShutdownReembed();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }

            // if the presentation has been saved before embed, save the presentation
            if (saved == Office.MsoTriState.msoTrue)
            {
                Pres.Save();

                // delete the temp folder only when user has saved file
                string tempFolderPath = Path.GetTempPath() + TempFolderName;
                if (Directory.Exists(tempFolderPath))
                {
                    Directory.Delete(tempFolderPath, true);
                }
            }
        }

        private void SlideShowBeginHandler(PowerPoint.SlideShowWindow wn)
        {
            isInSlideShow = true;
        }

        private void SlideShowEndHandler(PowerPoint.Presentation presentation)
        {
            isInSlideShow = false;

            // force recording session ends
            if (recorderTaskPane.HasEvent())
            {
                recorderTaskPane.ForceStopEvent();
            }

            // enable slide show button
            recorderTaskPane.EnableSlideShow();

            // when leave the show, dispose the in-show control if we have one
            recorderTaskPane.DisposeInSlideControlBox();

            // if audio buffer is not empty, render the effects
            if (recorderTaskPane.AudioBuffer.Count != 0)
            {
                var slides = PowerPointPresentation.Slides.ToList();

                for (int i = 0; i < recorderTaskPane.AudioBuffer.Count; i++)
                {
                    if (recorderTaskPane.AudioBuffer[i].Count != 0)
                    {
                        foreach (var audio in recorderTaskPane.AudioBuffer[i])
                        {
                            audio.Item1.EmbedOnSlide(slides[i], audio.Item2);
                        }
                    }
                }
            }

            // clear the buffer after embed
            recorderTaskPane.AudioBuffer.Clear();

            // change back the slide range settings
            Application.ActivePresentation.SlideShowSettings.RangeType = PowerPoint.PpSlideShowRangeType.ppShowAll;
        }

        void UpdateRecorderPane(int count, int id)
        {
            // if the user has selected none or more than 1 slides, recorder pane should show nothing
            if (count != 1)
            {
                if (customTaskPane.Visible)
                {
                    recorderTaskPane.ClearDisplayLists();
                }
            }
            else
            {
                // initailize the current slide
                recorderTaskPane.InitializeAudioAndScript(PowerPointPresentation.CurrentSlide, null, false);

                // if the pane is shown, refresh the pane immediately
                if (customTaskPane.Visible)
                {
                    recorderTaskPane.UpdateLists(id);
                }
            }
        }

        void BreakRecorderEvents()
        {
            // TODO:
            // Slide change event will interrupt mci device behaviour before
            // the event raised. Now we discard the record, we may want to
            // take this record by some means.
            if (recorderTaskPane.HasEvent())
            {
                recorderTaskPane.ForceStopEvent();
            }
        }

        void PrepareMediaFiles(PowerPoint.Presentation Pres)
        {
            string presFullName = Pres.FullName;
            string presName = Pres.Name;
            string tempPath = Path.GetTempPath() + TempFolderName;
            string zipName = presName.Replace(".pptx", ".zip");
            string zipFullPath = tempPath + zipName;

            // if temp folder doesn't exist, create
            if (!Directory.Exists(tempPath))
            {
                Directory.CreateDirectory(tempPath);
            }
            else
            // else clear the folder
            {
                Directory.Delete(tempPath, true);
                Directory.CreateDirectory(tempPath);
            }

            // copy the file to temp folder and rename to zip
            try
            {
                File.Copy(presFullName, zipFullPath);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }

            // open the zip and extract media files to temp folder
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
                        nameDictForPastedShapes[uniqueName] = nameListForCopiedShapes[i-1];
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
                var slideRange = selection.SlideRange;
                var oriSlide = 0;

                foreach (var sld in slideRange)
                {
                    var oldID = copiedSlides[oriSlide].SlideID;
                    var newID = (sld as PowerPoint.Slide).SlideID;
                    
                    recorderTaskPane.CopySlideToSlide(oldID, newID);

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
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    copiedSlides.Clear();

                    foreach (var sld in selection.SlideRange)
                    {
                        var slide = sld as PowerPoint.Slide;

                        copiedSlides.Add(slide);
                    }
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
                    copiedShapes.Sort((PowerPoint.Shape x, PowerPoint.Shape y) =>
                    {
                        return x.Id - y.Id;
                    });
                }
            }
            catch
            {
                //TODO: log in ThisAddIn.cs
            }
        }

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
