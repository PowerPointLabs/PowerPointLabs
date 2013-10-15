using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PPSpotlight
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PPSpotlight.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void ZoomBtnClick(Office.IRibbonControl control)
        {
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            PowerPoint.Shape picture = null;
            PowerPoint.Shape selectedShape = null;

            foreach (PowerPoint.Shape shape in Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange)
            {
                if (((PowerPoint.Shape)shape).Type == Office.MsoShapeType.msoPicture)
                {
                    picture = (PowerPoint.Shape)shape;
                }
                else
                {
                    selectedShape = (PowerPoint.Shape)shape;
                }
            }
            //PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
         
            float centerX = selectedShape.Left + selectedShape.Width / 2;
            float centerY = selectedShape.Top + selectedShape.Height / 2;

            picture.Copy();
            PowerPoint.Shape duplicatePic = currentSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];

            duplicatePic.PictureFormat.CropLeft += selectedShape.Left - picture.Left;
            duplicatePic.PictureFormat.CropTop += selectedShape.Top - picture.Top;
            duplicatePic.PictureFormat.CropRight += (picture.Left + picture.Width) - (selectedShape.Left + selectedShape.Width);
            duplicatePic.PictureFormat.CropBottom += (picture.Top + picture.Height) - (selectedShape.Top + selectedShape.Height);

            selectedShape.Delete();
            duplicatePic.Cut();

            currentSlide.Duplicate();
            //Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.CustomLayouts[7]);
            PowerPoint.Slide addedSlide = GetNextSlide(currentSlide);
            
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            PowerPoint.Shape sh = addedSlide.Shapes.Paste()[1];

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation; 
            sh.Width *= 2.0f;
            sh.Left = centerX - sh.Width / 2;
            sh.Top = centerY - sh.Height / 2;
            if (sh.Left < 0)
                sh.Left = 0;
            else if (sh.Left + sh.Width > presentation.PageSetup.SlideWidth)
                sh.Left = presentation.PageSetup.SlideWidth - sh.Width;
            if (sh.Top < 0)
                sh.Top = 0;
            else if (sh.Top + sh.Height > presentation.PageSetup.SlideHeight)
                sh.Top = presentation.PageSetup.SlideHeight - sh.Height;


            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            PowerPoint.Effect zoomEffect = null;
            zoomEffect = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectFadedZoom, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            zoomEffect.Timing.Duration = 0.5f;
        }

        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            currentSlide.Duplicate();
            PowerPoint.Slide addedSlide = GetNextSlide(currentSlide);

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Shape rectangleShape = addedSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
            rectangleShape.Fill.ForeColor.RGB = 0x000000;
            rectangleShape.Fill.Transparency = 0.2f;
            rectangleShape.Line.Visible = Office.MsoTriState.msoFalse;
            rectangleShape.Name = "SpotlightShape1";

            PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            selectedShape.Copy();

            foreach (PowerPoint.Shape sh in addedSlide.Shapes)
            {
                if (sh.Name.Equals(selectedShape.Name))
                {
                    sh.Delete();
                }
            }
            PowerPoint.Shape newShape = addedSlide.Shapes.Paste()[1];
            newShape.Left = selectedShape.Left;
            newShape.Top = selectedShape.Top;
            newShape.Fill.ForeColor.RGB = 0xffffff;
            newShape.Line.Visible = Office.MsoTriState.msoFalse;
            newShape.Name = "SpotlightShape2";
            selectedShape.Delete();

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            String[] array = { "SpotlightShape1", "SpotlightShape2" };
            PowerPoint.ShapeRange newRange = addedSlide.Shapes.Range(array);
            newRange.Select();

            PowerPoint.Selection currentSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            int count = currentSelection.ShapeRange.Count;
            currentSelection.Cut();

            PowerPoint.Shape pictureShape = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pictureShape.Left = 0;
            pictureShape.Top = 0;
            pictureShape.PictureFormat.TransparencyColor = 0xffffff;
            pictureShape.PictureFormat.TransparentBackground = Office.MsoTriState.msoTrue;
        }

        #endregion

        #region Helpers
        
        private PowerPoint.Slide GetCurrentSlide()
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            return Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
        }

        private PowerPoint.Slide GetNextSlide(PowerPoint.Slide currentSlide)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            int slideIndex = currentSlide.SlideIndex;
            return presentation.Slides[slideIndex + 1];
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
