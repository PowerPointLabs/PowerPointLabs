using System;
using System.IO;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Diagnostics;
using PPExtraEventHelper;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Deployment.Application;

namespace PowerPointLabs
{
    public partial class ThisAddIn
    {
        Ribbon1 ribbon;
        public ArrayList indicators = new ArrayList();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SetUpLogger();
            Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": PowerPointLabs Started");
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += new Microsoft.Office.Interop.PowerPoint.EApplication_NewPresentationEventHandler(ThisAddIn_NewPresentation);
            //((PowerPoint.EApplication_Event)this.Application).SlideShowBegin += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowBeginEventHandler(ThisAddIn_BeginSlideShow);
            //((PowerPoint.EApplication_Event)this.Application).SlideShowEnd += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowEndEventHandler(ThisAddIn_EndSlideShow);
            ((PowerPoint.EApplication_Event)this.Application).WindowSelectionChange += new Microsoft.Office.Interop.PowerPoint.EApplication_WindowSelectionChangeEventHandler(ThisAddIn_SelectionChanged);
            ((PowerPoint.EApplication_Event)this.Application).SlideSelectionChanged += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideSelectionChangedEventHandler(ThisAddIn_SlideSelectionChanged);
            //DisplayUpdateDetails();
            Application.SlideShowBegin += SlideShowBeginHandler;
            Application.SlideShowEnd += SlideShowEndHandler;
            PPMouse.Init(Application);
            SetupDoubleClickHandler();
            SetupTabActivateHandler();
        }

        void SetUpLogger()
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

        //void DisplayUpdateDetails()
        //{
        //    if (!ApplicationDeployment.IsNetworkDeployed)
        //        return;

        //    if (!ApplicationDeployment.CurrentDeployment.IsFirstRun)
        //        return;

        //    System.Windows.Forms.MessageBox.Show("PowerPointLabs has been updated recently.\nPlease visit http://powerpointlabs.info for more details", "Application Updated");
        //}

        void ThisAddIn_SelectionChanged(PowerPoint.Selection Sel)
        {
            ribbon.spotlightEnabled = false;
            ribbon.inSlideEnabled = false;
            ribbon.zoomButtonEnabled = false;
            if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape sh = Sel.ShapeRange[1];
                if (sh.Type == Office.MsoShapeType.msoAutoShape || sh.Type == Office.MsoShapeType.msoFreeform || sh.Type == Office.MsoShapeType.msoTextBox || sh.Type == Office.MsoShapeType.msoPlaceholder)
                {
                    ribbon.spotlightEnabled = true;
                }
                if ((sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle) || sh.Type == Office.MsoShapeType.msoPicture)
                {
                    ribbon.zoomButtonEnabled = true;
                }
                if (Sel.ShapeRange.Count > 1)
                {
                    ribbon.zoomButtonEnabled = false;
                    foreach (PowerPoint.Shape tempShape in Sel.ShapeRange)
                    {
                        if (sh.Type == tempShape.Type)
                            ribbon.inSlideEnabled = true;
                        if (sh.Type == Office.MsoShapeType.msoAutoShape && sh.AutoShapeType != tempShape.AutoShapeType)
                        {
                            ribbon.inSlideEnabled = false;
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

        void ThisAddIn_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            ribbon.addAutoMotionEnabled = true;
            ribbon.reloadAutoMotionEnabled = true;
            ribbon.reloadSpotlight = true;
            if (SldRange.Count != 1)
            {
                ribbon.addAutoMotionEnabled = false;
                ribbon.reloadAutoMotionEnabled = false;
                ribbon.reloadSpotlight = false;
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
                if (!((tmp.Name.Contains("PPSlideAnimated") && tmp.Name.Substring(0, 15).Equals("PPSlideAnimated"))
                    || ((tmp.Name.Contains("PPSlideStart") && tmp.Name.Substring(0, 12).Equals("PPSlideStart"))
                    && (next.Name.Contains("PPSlideAnimated") && next.Name.Substring(0, 15).Equals("PPSlideAnimated")))
                    || ((tmp.Name.Contains("PPSlideEnd") && tmp.Name.Substring(0, 10).Equals("PPSlideEnd"))
                    && (prev.Name.Contains("PPSlideAnimated") && prev.Name.Substring(0, 15).Equals("PPSlideAnimated")))
                    || ((tmp.Name.Contains("PPSlideMulti") && tmp.Name.Substring(0, 12).Equals("PPSlideMulti"))
                    && ((prev.Name.Contains("PPSlideAnimated") && prev.Name.Substring(0, 15).Equals("PPSlideAnimated"))
                    || (next.Name.Contains("PPSlideAnimated") && next.Name.Substring(0, 15).Equals("PPSlideAnimated"))))))
                    ribbon.reloadAutoMotionEnabled = false;
                if (!(tmp.Name.Contains("PPTLabsSpotlight")))
                    ribbon.reloadSpotlight = false;
            }
            ribbon.RefreshRibbonControl("AddAnimationButton");
            ribbon.RefreshRibbonControl("ReloadButton");
            ribbon.RefreshRibbonControl("ReloadSpotlightButton");
        }

        //void ThisAddIn_BeginSlideShow(PowerPoint.SlideShowWindow Wn)
        //{
        //    PowerPoint.Presentation pres = Wn.Presentation;
        //    indicators.Clear();

        //    foreach (PowerPoint.Slide sl in pres.Slides)
        //    {
        //        if (sl.Name.Contains("PPSlide") && sl.Name.Substring(0, 7).Equals("PPSlide"))
        //        {
        //            foreach (PowerPoint.Shape sh in sl.Shapes)
        //            {
        //                if (sh.Name.Contains("PPIndicator"))
        //                {
        //                    sh.Visible = Office.MsoTriState.msoFalse;
        //                    indicators.Add(sh);
        //                }
        //            }
        //        }
        //    }
        //}

        //void ThisAddIn_EndSlideShow(PowerPoint.Presentation Pres)
        //{
        //    foreach (PowerPoint.Shape sh in indicators)
        //    {
        //        sh.Visible = Office.MsoTriState.msoTrue;
        //    }
        //}

        void ThisAddIn_NewPresentation(PowerPoint.Presentation Pres)
        {
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

        private void SlideShowBeginHandler(PowerPoint.SlideShowWindow wn)
        {
            isInSlideShow = true;
        }

        private void SlideShowEndHandler(PowerPoint.Presentation presentation)
        {
            isInSlideShow = false;
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
            const int CommandOpenBackgroundFormat = 0x8F;
            var selectedShapes = selection.ShapeRange;
            Native.SendMessage(
                Process.GetCurrentProcess().MainWindowHandle,
                (uint)Native.Message.WM_COMMAND,
                new IntPtr(CommandOpenBackgroundFormat),
                IntPtr.Zero
                );
            if (!isInSlideShow)
            {
                selectedShapes.Select();
            }
        }

        //For office 2010 (in office 2013, this method has bad user exp)
        //Use hotkey (Alt - H - O) to activate Property window
        private void OpenPropertyWindowForOffice10()
        {
            try
            {
                string Shortcut_Alt_H_O = "%ho";
                if (eventHook == IntPtr.Zero)
                {
                    //Check whether Home tab is enabled or not
                    eventHook = Native.SetWinEventHook(
                        (uint)Native.Event.EVENT_SYSTEM_MENUEND,
                        (uint)Native.Event.EVENT_OBJECT_CREATE,
                        IntPtr.Zero,
                        TabActivate,
                        (uint)Process.GetCurrentProcess().Id,
                        0,
                        0);
                }
                SendKeys.Send(Shortcut_Alt_H_O);
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
