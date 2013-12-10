using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml.Linq;
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
            ((PowerPoint.EApplication_Event)this.Application).NewPresentation += new Microsoft.Office.Interop.PowerPoint.EApplication_NewPresentationEventHandler(ThisAddIn_NewPresentation);
            //((PowerPoint.EApplication_Event)this.Application).SlideShowBegin += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowBeginEventHandler(ThisAddIn_BeginSlideShow);
            //((PowerPoint.EApplication_Event)this.Application).SlideShowEnd += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideShowEndEventHandler(ThisAddIn_EndSlideShow);
            ((PowerPoint.EApplication_Event)this.Application).WindowSelectionChange += new Microsoft.Office.Interop.PowerPoint.EApplication_WindowSelectionChangeEventHandler(ThisAddIn_SelectionChanged);
            ((PowerPoint.EApplication_Event)this.Application).SlideSelectionChanged += new Microsoft.Office.Interop.PowerPoint.EApplication_SlideSelectionChangedEventHandler(ThisAddIn_SlideSelectionChanged);
            DisplayUpdateDetails();
        }

        void DisplayUpdateDetails()
        {
            if (!ApplicationDeployment.IsNetworkDeployed)
                return;

            if (!ApplicationDeployment.CurrentDeployment.IsFirstRun)
                return;

            System.Windows.Forms.MessageBox.Show("PowerPointLabs has been updated recently.\nPlease visit http://powerpointlabs.info for more details", "Application Updated");
        }

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
                if (!(tmp.Name.Contains("PPSlide") && tmp.Name.Substring(0, 7).Equals("PPSlide")))
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
