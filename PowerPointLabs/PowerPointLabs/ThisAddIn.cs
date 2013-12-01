using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

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
