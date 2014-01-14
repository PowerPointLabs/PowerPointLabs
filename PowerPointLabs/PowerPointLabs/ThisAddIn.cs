using System;
using System.IO;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Diagnostics;
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
            if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape sh = Sel.ShapeRange[1];
                if (sh.Type == Office.MsoShapeType.msoAutoShape || sh.Type == Office.MsoShapeType.msoFreeform || sh.Type == Office.MsoShapeType.msoTextBox || sh.Type == Office.MsoShapeType.msoPlaceholder)
                {
                    ribbon.spotlightEnabled = true;
                }
            }

            ribbon.RefreshRibbonControl("AddSpotlightButton");
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

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
