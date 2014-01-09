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


namespace PowerPointLabs
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public float defaultSoftEdges = 10;
        public float defaultDuration = 0.5f;
        public float defaultTransparency = 0.7f;
        public bool startUp = false;
        public bool spotlightEnabled = false;
        public bool addAutoMotionEnabled = true;
        public bool reloadAutoMotionEnabled = true;
        public bool reloadSpotlight = true;
        public Dictionary<String, float> softEdgesMapping = new Dictionary<string,float>
        {
            {"None", 0},
            {"1 Point", 1},
            {"2.5 Points", 2.5f},
            {"5 Points", 5},
            {"10 Points", 10},
            {"25 Points", 25},
            {"50 Points", 50}
        };

        public Dictionary<String, PowerPoint.Shape> spotlightShapeMapping = new Dictionary<string,PowerPoint.Shape>();
        public Dictionary<String, PowerPoint.Slide> spotlightSlideMapping = new Dictionary<string, PowerPoint.Slide>();

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointLabs.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void RefreshRibbonControl(String controlID)
        {
            ribbon.InvalidateControl(controlID);
        }

        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            //Get References of current and next slides
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            if (currentSlide != null && currentSlide.SlideIndex != presentation.Slides.Count)
            {
                PowerPoint.Slide nextSlide = GetNextSlide(currentSlide);
                AddCompleteAutoMotion(currentSlide, nextSlide);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select the correct slide", "Unable to Add Animations");
            }
        }

        public void ReloadButtonClick(Office.IRibbonControl control)
        {
            PowerPoint.Slide tempSlide = GetCurrentSlide();
            if (tempSlide.Name.Contains("PPSlideAnimated") && tempSlide.Name.Substring(0, 15).Equals("PPSlideAnimated"))
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide nextSlide = presentation.Slides[tempSlide.SlideIndex + 1];
                PowerPoint.Slide currentSlide = presentation.Slides[tempSlide.SlideIndex - 1];
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.SlideIndex);
                tempSlide.Delete();

                AddCompleteAutoMotion(currentSlide, nextSlide);
            }
            else if (tempSlide.Name.Contains("PPSlideStart") && tempSlide.Name.Substring(0, 12).Equals("PPSlideStart"))
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide animatedSlide = presentation.Slides[tempSlide.SlideIndex + 1];
                PowerPoint.Slide finalSlide = presentation.Slides[tempSlide.SlideIndex + 2];
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);
                animatedSlide.Delete();

                AddCompleteAutoMotion(tempSlide, finalSlide);
            }
            else if (tempSlide.Name.Contains("PPSlideEnd") && tempSlide.Name.Substring(0, 10).Equals("PPSlideEnd"))
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide animatedSlide = presentation.Slides[tempSlide.SlideIndex - 1];
                PowerPoint.Slide firstSlide = presentation.Slides[tempSlide.SlideIndex - 2];
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);
                animatedSlide.Delete();

                AddCompleteAutoMotion(firstSlide, tempSlide);
            }
            else if (tempSlide.Name.Contains("PPSlideMulti") && tempSlide.Name.Substring(0, 12).Equals("PPSlideMulti"))
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide startSlide1 = tempSlide;
                PowerPoint.Slide animatedSlide1 = tempSlide;
                PowerPoint.Slide animatedSlide2 = tempSlide;
                PowerPoint.Slide endSlide2 = tempSlide;
                if (tempSlide.SlideIndex > 2)
                {
                    animatedSlide1 = presentation.Slides[tempSlide.SlideIndex - 1];
                    startSlide1 = presentation.Slides[tempSlide.SlideIndex - 2];
                    if (animatedSlide1.Name.Contains("PPSlideAnimated") && startSlide1.Name.Contains("PPSlideStart"))
                    {
                        Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);
                        animatedSlide1.Delete();
                        AddCompleteAutoMotion(startSlide1, tempSlide);
                    }
                }

                if (tempSlide.SlideIndex < presentation.Slides.Count - 1)
                {
                    animatedSlide2 = presentation.Slides[tempSlide.SlideIndex + 1];
                    endSlide2 = presentation.Slides[tempSlide.SlideIndex + 2];
                    if (animatedSlide2.Name.Contains("PPSlideAnimated") && endSlide2.Name.Contains("PPSlideEnd"))
                    {
                        Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);
                        animatedSlide2.Delete();
                        AddCompleteAutoMotion(tempSlide, endSlide2);
                    }
                }    
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The current slide was not added by PPTLabs AutoMotion", "Error");
            }
        }

        public void AboutButtonClick(Office.IRibbonControl control)
        {
            //AboutForm form = new AboutForm();
            //form.Show();
            System.Windows.Forms.MessageBox.Show("          PowerPointLabs Plugin Version 1.2.0 [Release date: 8 Jan 2014]\n     Developed at School of Computing, National University of Singapore.\n        For more information, visit our website http://PowerPointLabs.info", "About PowerPointLabs");
        }

        public void HelpButtonClick(Office.IRibbonControl control)
        {
            string myURL = "http://powerpointlabs.info/docs.html";
            System.Diagnostics.Process.Start(myURL);
        }

        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
            string myURL = "http://powerpointlabs.info/contact.html";
            System.Diagnostics.Process.Start(myURL);
        }

        public void HighlightBulletsButtonClick(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("This feature is coming soon!                ", "Coming Soon");
        }

        public void AddZoomButtonClick(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("This feature is coming soon!                ", "Coming Soon");
        }

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("This feature is coming soon!                ", "Coming Soon");
        }

        //public void AboutButtonClickSpotlight(Office.IRibbonControl control)
        //{
        //    AboutSpotlight form = new AboutSpotlight();
        //    form.Show();
        //}

        //Dropdown Callbacks
        //public int OnGetItemCount(Office.IRibbonControl control)
        //{
        //    return effectMapping.Count;
        //}

        //public String OnGetItemLabel(Office.IRibbonControl control, int index)
        //{
        //    String[] keys = effectMapping.Keys.ToArray();
        //    return keys[index];
        //}

        //public void OnSelectItem(Office.IRibbonControl control, String selectedId, int selectedIndex)
        //{
        //    String[] keys = effectMapping.Keys.ToArray();
        //    defaultEffect = effectMapping[keys[selectedIndex]];
        //}

        public System.Drawing.Bitmap GetAddAnimationImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.AddAnimation);
        }

        public System.Drawing.Bitmap GetReloadAnimationImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.ReloadAnimation);
        }

        public System.Drawing.Bitmap GetSpotlightImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.Spotlight);
        }

        public System.Drawing.Bitmap GetReloadSpotlightImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.ReloadSpotlight);
        }

        public System.Drawing.Bitmap GetHighlightBulletsImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.Bullets);
        }
        public System.Drawing.Bitmap GetZoomImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.Zoom);
        }
        public System.Drawing.Bitmap GetCropShapeImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.CutOutShape);
        }

        public System.Drawing.Bitmap GetAboutImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.About);
        }
        public System.Drawing.Bitmap GetHelpImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.Help);
        }
        public System.Drawing.Bitmap GetFeedbackImage(Office.IRibbonControl control)
        {
            return new System.Drawing.Bitmap(Properties.Resources.Feedback);
        }
        //Duration Callbacks
        public void OnChangeDuration(Office.IRibbonControl control, String text)
        {
            if (text == "")
                defaultDuration = 0.01f;
            else
            {
                float enteredValue = float.Parse(text);
                if (enteredValue < 0.01)
                    defaultDuration = 0.01f;
                else if (enteredValue > 59.0)
                    defaultDuration = 59.0f;
                else
                    defaultDuration = enteredValue;
            }
            ribbon.InvalidateControl("animationDurationOption");
        }

        public String OnGetText(Office.IRibbonControl control)
        {
            return defaultDuration.ToString();
        }

        //Transparency Callbacks
        public void OnChangeTransparency(Office.IRibbonControl control, String text)
        {
            if (text.Contains('%'))
            {
                text = text.Substring(0, text.IndexOf('%'));    
            }
            float result;
            if (float.TryParse(text, out result))
            {
                if (result > 0 && result <= 100)
                {
                    defaultTransparency = result;
                    defaultTransparency /= 100;
                }
            }
            ribbon.InvalidateControl("spotlightTransparency");
        }

        public String OnGetTransparency(Office.IRibbonControl control)
        {
            return (defaultTransparency*100).ToString() + "%";
        }

        public int OnGetItemCountSpotlight(Office.IRibbonControl control)
        {
            return softEdgesMapping.Count;
        }

        public String OnGetItemLabelSpotlight(Office.IRibbonControl control, int index)
        {
            String[] keys = softEdgesMapping.Keys.ToArray();
            return keys[index];
        }

        public void OnSelectItemSpotlight(Office.IRibbonControl control, String selectedId, int selectedIndex)
        {
            String[] keys = softEdgesMapping.Keys.ToArray();
            defaultSoftEdges = softEdgesMapping[keys[selectedIndex]];
        }

        public int OnGetSelectedItemIndexSpotlight(Office.IRibbonControl control)
        {
            float[] values = softEdgesMapping.Values.ToArray();
            return Array.IndexOf(values, defaultSoftEdges);
        }

        public bool OnGetEnabledSpotlight(Office.IRibbonControl control)
        {
            return spotlightEnabled;
        }

        public bool OnGetEnabledReloadSpotlight(Office.IRibbonControl control)
        {
            return reloadSpotlight;
        }

        public bool OnGetEnabledAddAutoMotion(Office.IRibbonControl control)
        {
            return addAutoMotionEnabled;
        }

        public bool OnGetEnabledReloadAutoMotion(Office.IRibbonControl control)
        {
            return reloadAutoMotionEnabled;
        }

        //Edit Name Callbacks
        public void NameEditBtnClick(Office.IRibbonControl control)
        {
            PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            Form1 editForm = new Form1(this, selectedShape.Name);
            editForm.Show();
        }

        public void nameEdited(String newName)
        {
            PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            selectedShape.Name = newName;
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

            if (picture == null || selectedShape == null)
            {
                System.Windows.Forms.MessageBox.Show("Unable to add zoom animation for selected objects", "Error");
            }
            else
            {
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

                //currentSlide.Duplicate();
                Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.CustomLayouts[7]);
                PowerPoint.Slide addedSlide = GetNextSlide(currentSlide);

                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
                PowerPoint.Shape sh = addedSlide.Shapes.Paste()[1];
                sh.Left = centerX - (sh.Width / 2);
                sh.Top = centerY - (sh.Height / 2);

                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Effect effectMotion = null;
                PowerPoint.Effect effectResize = null;
                PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
                float finalX = (presentation.PageSetup.SlideWidth / 2);
                float initialX = (sh.Left + (sh.Width) / 2);
                float finalY = (presentation.PageSetup.SlideHeight / 2);
                float initialY = (sh.Top + (sh.Height) / 2);

                effectMotion = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                effectMotion.Timing.Duration = defaultDuration;
                motion.MotionEffect.Path = "M 0 0 C " + ((finalX - initialX) / 2) / presentation.PageSetup.SlideWidth + " " + ((finalY - initialY) / 2) / presentation.PageSetup.SlideHeight + " " + ((finalX - initialX) / 2) / presentation.PageSetup.SlideWidth + " " + ((finalY - initialY) / 2) / presentation.PageSetup.SlideHeight + " " + (finalX - initialX) / presentation.PageSetup.SlideWidth + " " + (finalY - initialY) / presentation.PageSetup.SlideHeight + " E";
                    

                float finalWidth = presentation.PageSetup.SlideWidth;
                float initialWidth = sh.Width;
                float finalHeight = presentation.PageSetup.SlideHeight;
                float initialHeight = sh.Height;

                effectResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];

                effectResize.Timing.Duration = defaultDuration;
                resize.ScaleEffect.ToX = (finalWidth / initialWidth) * 100;
                resize.ScaleEffect.ToY = (finalHeight / initialHeight) * 100;
                //sh.Width = presentation.PageSetup.SlideWidth;
                //sh.Height = presentation.PageSetup.SlideHeight;
                //sh.Left = (presentation.PageSetup.SlideWidth / 2) - (sh.Width / 2);
                //sh.Top = (presentation.PageSetup.SlideHeight / 2) - (sh.Height / 2);
                //sh.Width *= 2.0f;
                //sh.Left = centerX - sh.Width / 2;
                //sh.Top = centerY - sh.Height / 2;
                //if (sh.Left < 0)
                //    sh.Left = 0;
                //else if (sh.Left + sh.Width > presentation.PageSetup.SlideWidth)
                //    sh.Left = presentation.PageSetup.SlideWidth - sh.Width;
                //if (sh.Top < 0)
                //    sh.Top = 0;
                //else if (sh.Top + sh.Height > presentation.PageSetup.SlideHeight)
                //    sh.Top = presentation.PageSetup.SlideHeight - sh.Height;


                //PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
                //PowerPoint.Effect zoomEffect = null;
                //zoomEffect = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectFadedZoom, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                //zoomEffect.Timing.Duration = 0.5f;
                AddAckSlide();
            }
        }

        public void ReloadSpotlightButtonClick(Office.IRibbonControl control)
        {
            PowerPoint.Slide tempSlide = GetCurrentSlide();
            PowerPoint.Shape shape1 = null;
            List<PowerPoint.Shape> spotlightShapes = new List<PowerPoint.Shape>();
            if (tempSlide.Name.Contains("PPTLabsSpotlight")) //&& tempSlide.Name.Substring(0, 14).Equals("PPTLabsSpotlight")
            {
                foreach (PowerPoint.Shape sh in tempSlide.Shapes)
                {
                    if (sh.Name.Equals("SpotlightShape1"))
                    {
                        shape1 = sh;
                    }
                    else if (sh.Name.Contains("SpotlightShape"))
                    {
                        spotlightShapes.Add(sh);
                    }
                }

                if (shape1 == null || spotlightShapes.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("The current slide cannot be reloaded", "Error");
                }
                else
                {
                    shape1.Delete();

                    foreach (PowerPoint.Shape sh in spotlightShapes)
                    {
                        sh.Visible = Office.MsoTriState.msoTrue;

                        PowerPoint.Shape duplicateShape = sh.Duplicate()[1];
                        duplicateShape.Visible = Office.MsoTriState.msoFalse;
                        duplicateShape.Left = sh.Left;
                        duplicateShape.Top = sh.Top;
                    }
                    
                    AddSpotlightEffect(tempSlide, spotlightShapes);
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The current slide was not added by PPTLabs Spotlight", "Error");
            }
        }

        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count == 0)
                return;

            PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

            currentSlide.Duplicate();
            PowerPoint.Slide addedSlide = GetNextSlide(currentSlide);
            addedSlide.Name = "PPTLabsSpotlight" + GetTimestamp(DateTime.Now);
            int counter = 2;
            List<PowerPoint.Shape> spotlightShapes = new List<PowerPoint.Shape>();

            foreach (PowerPoint.Shape spotShape in selectedShapes)
            {
                foreach (PowerPoint.Shape copyShape in addedSlide.Shapes)
                {
                    if (copyShape.Name.Equals(spotShape.Name))
                    {
                        copyShape.Delete();
                    }
                }
                spotShape.Copy();
                PowerPoint.Shape spotlightShape = addedSlide.Shapes.Paste()[1];

                if (spotShape.Left < 0)
                {
                    spotlightShape.Left = 0;
                    spotlightShape.Width = spotShape.Width - (0.0f - spotShape.Left);
                }
                else
                    spotlightShape.Left = spotShape.Left;

                if (spotShape.Left + spotShape.Width > Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth)
                    spotlightShape.Width = (Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth - spotlightShape.Left);

                if (spotShape.Top < 0)
                {
                    spotlightShape.Top = 0;
                    spotlightShape.Height = spotShape.Height - (0.0f - spotShape.Top);
                }
                else
                    spotlightShape.Top = spotShape.Top;

                if (spotShape.Top + spotShape.Height > Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight)
                    spotlightShape.Height = (Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight - spotlightShape.Top);

                spotlightShape.Fill.ForeColor.RGB = 0xffffff;
                spotlightShape.Line.Visible = Office.MsoTriState.msoFalse;
                spotlightShape.Name = "SpotlightShape" + counter;
                counter++;

                PowerPoint.Shape duplicateShape = spotlightShape.Duplicate()[1];
                duplicateShape.Visible = Office.MsoTriState.msoFalse;
                duplicateShape.Left = spotlightShape.Left;
                duplicateShape.Top = spotlightShape.Top;

                spotlightShapes.Add(spotlightShape);
                spotShape.Delete();
            }

            AddSpotlightEffect(addedSlide, spotlightShapes);
            AddAckSlide();
        }

        #endregion

        #region Helpers

        private void AddSpotlightEffect(PowerPoint.Slide addedSlide,List<PowerPoint.Shape> spotlightShapes)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Shape rectangleShape = addedSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
            rectangleShape.Fill.ForeColor.RGB = 0x000000;
            rectangleShape.Fill.Transparency = defaultTransparency;
            rectangleShape.Line.Visible = Office.MsoTriState.msoFalse;
            rectangleShape.Name = "SpotlightShape1";
            rectangleShape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
            
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            List<String> shapeNames = new List<String>();
            shapeNames.Add("SpotlightShape1");
            foreach (PowerPoint.Shape sh in spotlightShapes)
            {
                shapeNames.Add(sh.Name);
            }
            String[] array = shapeNames.ToArray();
            PowerPoint.ShapeRange newRange = addedSlide.Shapes.Range(array);
            newRange.Select();

            PowerPoint.Selection currentSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            int count = currentSelection.ShapeRange.Count;
            currentSelection.Cut();

            PowerPoint.Shape pictureShape = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pictureShape.PictureFormat.TransparencyColor = 0xffffff;
            pictureShape.PictureFormat.TransparentBackground = Office.MsoTriState.msoTrue;
            pictureShape.Left = 0.0f - defaultSoftEdges;
            pictureShape.Top = 0.0f - defaultSoftEdges;
            pictureShape.LockAspectRatio = Office.MsoTriState.msoFalse;
            pictureShape.Width = presentation.PageSetup.SlideWidth + (2.0f * defaultSoftEdges);
            pictureShape.Height = presentation.PageSetup.SlideHeight + (2.0f * defaultSoftEdges);
            pictureShape.SoftEdge.Radius = defaultSoftEdges;
            pictureShape.Name = "SpotlightShape1";
        }

        private void AddCompleteAutoMotion(PowerPoint.Slide currentSlide, PowerPoint.Slide nextSlide)
        {
            //Create containers to store information on matching shapes
            PowerPoint.Shape[] shapes1;
            PowerPoint.Shape[] shapes2;
            int[] shapeIDs;

            if (GetMatchingShapeDetails(currentSlide, nextSlide, out shapes1, out shapes2, out shapeIDs))
            {
                //If an identical object exists
                PowerPoint.Slide newSlide = PrepareAnimatedSlide(currentSlide, shapeIDs);
                AddAnimationsToShapes(newSlide, shapes1, shapes2, shapeIDs);
                //this.ribbon.ActivateTabMso("TabAnimations");
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
            }
            else
            {
                //Display error message
                System.Windows.Forms.MessageBox.Show("No matching Shapes were found on the next slide", "Animation Not Added");
            }
        }

        private void AddAnimationsToShapes(PowerPoint.Slide newSlide, PowerPoint.Shape[] shapes1, PowerPoint.Shape[] shapes2, int[] shapeIDs)
        {
            int count = 0;
            bool fadeFlag = false;
            PowerPoint.Sequence sequence = newSlide.TimeLine.MainSequence;
            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectRotate = null;
            PowerPoint.Effect effectFontResize = null;
            PowerPoint.Effect effectFade = null;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.MsoAnimTriggerType trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = newSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;
            int indicatorID = indicatorShape.Id;

            foreach (PowerPoint.Shape sh in newSlide.Shapes)
            {
                if (!shapeIDs.Contains(sh.Id) && sh.Id != indicatorID)
                {
                    //sh.Delete();
                    DeleteShapeAnnimations(newSlide, sh);
                    effectFade = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = defaultDuration;
                    fadeFlag = true;
                }
            }

            //Add animation effects to the duplicated objects
            foreach (PowerPoint.Shape sh in newSlide.Shapes)
            {
                if (shapeIDs.Contains(sh.Id))
                {
                    if (count < shapeIDs.Count() && sh.Id == shapeIDs[count])
                    {
                        DeleteShapeAnnimations(newSlide, sh);
                        trigger = (count == 0 && fadeFlag) ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious : PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        float finalX = (shapes2[count].Left + (shapes2[count].Width) / 2);
                        float initialX = (shapes1[count].Left + (shapes1[count].Width) / 2);
                        float finalY = (shapes2[count].Top + (shapes2[count].Height) / 2);
                        float initialY = (shapes1[count].Top + (shapes1[count].Height) / 2);

                        float finalRotation = shapes2[count].Rotation;
                        float initialRotation = shapes1[count].Rotation;

                        float finalWidth = shapes2[count].Width;
                        float initialWidth = shapes1[count].Width;
                        float finalHeight = shapes2[count].Height;
                        float initialHeight = shapes1[count].Height;
                        float finalFont = 0.0f;
                        float initialFont = 0.0f;

                        if (sh.HasTextFrame == Office.MsoTriState.msoTrue && (sh.TextFrame.HasText == Office.MsoTriState.msoTriStateMixed || sh.TextFrame.HasText == Office.MsoTriState.msoTrue) && sh.TextFrame.TextRange.Font.Size != shapes2[count].TextFrame.TextRange.Font.Size)
                        {
                            finalFont = shapes2[count].TextFrame.TextRange.Font.Size;
                            initialFont = shapes1[count].TextFrame.TextRange.Font.Size;  
                        }

                        if (finalHeight != initialHeight || finalWidth != initialWidth || finalFont != initialFont)
                        {
                            float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / 50.0f;
                            float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / 50.0f;
                            float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / 50.0f;
                            float incrementLeft = (finalX - initialX) / 50.0f;
                            float incrementTop = (finalY - initialY) / 50.0f;
                            float incrementFont = (finalFont - initialFont) / 50.0f;

                            PowerPoint.Shape lastShape = sh;
                            for (int i = 1; i <= 50; i++)
                            {
                                PowerPoint.Shape dupShape = sh.Duplicate()[1];
                                if (i != 1)
                                    sequence[sequence.Count].Delete();

                                dupShape.Left = sh.Left;
                                dupShape.Top = sh.Top;
                                dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                dupShape.Rotation += (incrementRotation * i);
                                dupShape.Left += (incrementLeft * i);
                                dupShape.Top += (incrementTop * i);

                                if (incrementFont != 0.0f)
                                {
                                    dupShape.TextFrame.TextRange.Font.Size += (incrementFont * i);
                                }

                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                //appear.Timing.Duration = 0.005f;
                                appear.Timing.TriggerDelayTime = ((defaultDuration / 50) * i);

                                PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                disappear.Exit = Office.MsoTriState.msoTrue;
                                //disappear.Timing.Duration = 0.005f;
                                disappear.Timing.TriggerDelayTime = ((defaultDuration / 50) * i);

                                lastShape = dupShape;
                            }
                        }
                        else 
                        {
                            //Motion Effect
                            if ((finalX != initialX) || (finalY != initialY))
                            {
                                effectMotion = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                                //PowerPoint.AnimationBehavior motion = effectMotion.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);
                                PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                                effectMotion.Timing.Duration = defaultDuration;
                                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                                //effectMotion.EffectParameters.Relative = Office.MsoTriState.msoTrue;
                                //motion.MotionEffect.FromX = initialX / presentation.PageSetup.SlideWidth * 100;
                                //motion.MotionEffect.FromY = initialY / presentation.PageSetup.SlideHeight * 100;
                                //motion.MotionEffect.ToX = (finalX - initialX) / presentation.PageSetup.SlideWidth * 100;
                                //motion.MotionEffect.ToY = (finalY - initialY) / presentation.PageSetup.SlideHeight * 100;

                                //Create VML path for the motion path
                                //This path needs to be a curved path to allow the user to edit points
                                float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                                float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                                float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                                float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                                float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                                float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;
                                motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                            }

                            //Resize Effect
                            //if (sh.Type != Office.MsoShapeType.msoPlaceholder && sh.Type != Office.MsoShapeType.msoTextBox)
                            //{
                            //    if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
                            //    {
                            //        sh.LockAspectRatio = Office.MsoTriState.msoTrue;
                            //        effectResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            //        //PowerPoint.AnimationBehavior resize = effectResize.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeScale);
                            //        PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];

                            //        //float rotCos = (float)Math.Cos(degToRad(sh.Rotation));
                            //        //float rotSin = (float)Math.Sin(degToRad(sh.Rotation));

                            //        effectResize.Timing.Duration = defaultDuration;
                            //        //sh.ScaleWidth((finalWidth / initialWidth), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                            //        //sh.ScaleHeight((finalHeight / initialHeight), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);

                            //        resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                            //        resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                            //        //resize.ScaleEffect.ByX = (((finalWidth / initialWidth) * Math.Abs(rotCos)) + ((finalHeight / initialHeight) * Math.Abs(rotSin))) * 100;
                            //        //resize.ScaleEffect.ByY = (((finalWidth / initialWidth) * Math.Abs(rotSin)) + ((finalHeight / initialHeight) * Math.Abs(rotCos))) * 100;
                            //        trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                            //    }
                            //}
                            //if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame.HasText == Office.MsoTriState.msoTrue && sh.TextFrame.TextRange.Font.Size != shapes2[count].TextFrame.TextRange.Font.Size)
                            //{
                            //    sh.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                            //    effectFontResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontSize, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            //    effectFontResize.Timing.Duration = defaultDuration;
                            //    PowerPoint.AnimationBehavior resizeFont = effectFontResize.Behaviors[1];
                            //    resizeFont.PropertyEffect.To = shapes2[count].TextFrame.TextRange.Font.Size / shapes1[count].TextFrame.TextRange.Font.Size;
                            //    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                            //}

                            //Rotation Effect
                            if (finalRotation != initialRotation)
                            {
                                effectRotate = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                                PowerPoint.AnimationBehavior rotate = effectRotate.Behaviors[1];
                                effectRotate.Timing.Duration = defaultDuration;
                                effectRotate.EffectParameters.Amount = GetMinimumRotation(initialRotation, finalRotation);
                                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                            }
                        }
                        count++;
                    }
                }
            }
        }

        private PowerPoint.Slide PrepareAnimatedSlide(PowerPoint.Slide currentSlide, int[] shapeIDs)
        {
            //Duplicate current slide
            currentSlide.Duplicate();
            if (currentSlide.Name.Contains("PPSlideEnd") || currentSlide.Name.Contains("PPSlideMulti"))
                currentSlide.Name = "PPSlideMulti" + GetTimestamp(DateTime.Now);
            else
                currentSlide.Name = "PPSlideStart" + GetTimestamp(DateTime.Now);
            //Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, currentSlide.CustomLayout);

            //Store reference of new slide
            PowerPoint.Slide newSlide = GetNextSlide(currentSlide);
            newSlide.Name = "PPSlideAnimated" + GetTimestamp(DateTime.Now);

            //Go to new slide
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(newSlide.SlideIndex);

            //Delete non-identical shapes
            //foreach (PowerPoint.Shape sh in newSlide.Shapes)
            //{
            //    if (!shapeIDs.Contains(sh.Id))
            //    {
            //        //sh.Delete();
            //        fadeEffect = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            //        fadeEffect.Timing.Duration = defaultDuration;
            //    }
            //}

            //Manage Slide Transitions
            newSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            //newSlide.SlideShowTransition.Duration = defaultDuration;
            newSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            newSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            newSlide.SlideShowTransition.AdvanceTime = 0;

            PowerPoint.Slide nextSlide = GetNextSlide(newSlide);
            nextSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            if (nextSlide.Name.Contains("PPSlideStart") || nextSlide.Name.Contains("PPSlideMulti"))
                nextSlide.Name = "PPSlideMulti" + GetTimestamp(DateTime.Now);
            else
                nextSlide.Name = "PPSlideEnd" + GetTimestamp(DateTime.Now);

            AddAckSlide();

            return newSlide;
        }

        private void AddAckSlide()
        {
            PowerPoint.Slide tempSlide = (Globals.ThisAddIn.Application.ActivePresentation.Slides[Globals.ThisAddIn.Application.ActivePresentation.Slides.Count]);
            if (!(tempSlide.Name.Contains("PPAck") && tempSlide.Name.Substring(0, 5).Equals("PPAck")))
            {
                PowerPoint.Slide ackSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides.Count + 1, GetCurrentSlide().CustomLayout);
                //Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(ackSlide.SlideIndex);
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                String tempFileName = Path.GetTempFileName();
                Properties.Resources.Acknowledgement.Save(tempFileName);
                float width = presentation.PageSetup.SlideWidth * 0.858f;
                float height = presentation.PageSetup.SlideHeight * (5.33f / 7.5f);
                PowerPoint.Shape ackShape = ackSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, ((presentation.PageSetup.SlideWidth - width) / 2), ((presentation.PageSetup.SlideHeight - height) / 2), width, height);
                ackSlide.SlideShowTransition.Hidden = Office.MsoTriState.msoTrue;
                ackSlide.Name = "PPAck" + GetTimestamp(DateTime.Now);
            }
        }

        private bool GetMatchingShapeDetails(PowerPoint.Slide currentSlide, PowerPoint.Slide nextSlide, out PowerPoint.Shape[] shapes1, out PowerPoint.Shape[] shapes2, out int[] shapeIDs)
        {
            bool flag = false;
            int counter = 0;

            shapes1 = new PowerPoint.Shape[currentSlide.Shapes.Count];
            shapes2 = new PowerPoint.Shape[currentSlide.Shapes.Count];
            shapeIDs = new int[currentSlide.Shapes.Count];

            foreach (PowerPoint.Shape sh1 in currentSlide.Shapes)
            {
                foreach (PowerPoint.Shape sh2 in nextSlide.Shapes)
                {
                    if (haveSameNames(sh1, sh2))
                    {
                        flag = true;
                        //if (sh1.Type == Office.MsoShapeType.msoPlaceholder && sh1.HasTextFrame == Office.MsoTriState.msoTrue)
                        //{
                        //    sh1.TextFrame.TextRange.Text.Trim();
                        //    sh1.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        //}
                        //if (sh2.Type == Office.MsoShapeType.msoPlaceholder && sh2.HasTextFrame == Office.MsoTriState.msoTrue)
                        //{
                        //    sh2.TextFrame.TextRange.Text.Trim();
                        //    sh2.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        //}

                        shapes1[counter] = sh1;
                        shapes2[counter] = sh2;
                        shapeIDs[counter] = sh1.Id;
                        counter++;
                        break;
                    }
                }
            }
            return flag;
        }

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

        private bool haveSameNames(PowerPoint.Shape sh1, PowerPoint.Shape sh2)
        {
            String name1 = sh1.Name;
            String name2 = sh2.Name;

            if (name1.ToUpper().CompareTo(name2.ToUpper()) == 0)
                return true;
            else
                return false;
        }

        public String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }

        private float GetMinimumRotation(float fromAngle, float toAngle)
        {
            fromAngle = Normalize(fromAngle);
            toAngle = Normalize(toAngle);

            float rotation1 = toAngle - fromAngle;
            float rotation2 = rotation1 == 0.0f ? 0.0f : Math.Abs(360.0f - Math.Abs(rotation1)) * (rotation1 / Math.Abs(rotation1)) * -1.0f;

            if (Math.Abs(rotation1) < Math.Abs(rotation2))
            {
                return rotation1;
            }
            else
            {
                return rotation2;
            }
        }

        private float Normalize(float i)
        {
            //find effective angle
            float d = Math.Abs(i) % 360.0f;

            if (i < 0)
            {
                //return positive equivalent
                return 360.0f - d;
            }
            else
            {
                return d;
            }
        }

        private void DeleteShapeAnnimations(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name)
                    effect.Delete();
            }

            PowerPoint.Slide nextSlide = GetNextSlide(slide);
            sequence = nextSlide.TimeLine.MainSequence;
            for (int x = sequence.Count; x >= 1; x--)
            {
                PowerPoint.Effect effect = sequence[x];
                if (effect.Shape.Name == shape.Name)
                    effect.Delete();
            }
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
