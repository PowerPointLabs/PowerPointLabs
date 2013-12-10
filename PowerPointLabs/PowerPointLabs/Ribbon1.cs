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

        public float defaultSoftEdges = 5;
        public float defaultDuration = 0.5f;
        public float defaultTransparency = 0.3f;
        public bool startUp = false;
        public bool spotlightEnabled = false;
        public Dictionary<String, float> softEdgesMapping = new Dictionary<string,float>
        {
            {"No Edges", 0},
            {"1 Point", 1},
            {"2.5 Points", 2.5f},
            {"5 Points", 5},
            {"10 Points", 10},
            {"25 Points", 25},
            {"50 Points", 50}
        };

        public Dictionary<String, PowerPoint.Shape> spotlightShapeMapping = new Dictionary<string,PowerPoint.Shape>();
        public Dictionary<String, PowerPoint.Slide> spotlightSlideMapping = new Dictionary<string, PowerPoint.Slide>();
        //public Dictionary<String, PowerPoint.MsoAnimEffect> effectMapping = new Dictionary<String, PowerPoint.MsoAnimEffect>
        //{ 
        //    {"Appear", PowerPoint.MsoAnimEffect.msoAnimEffectAppear},
        //    {"Arc Up", PowerPoint.MsoAnimEffect.msoAnimEffectArcUp},
        //    {"Ascend", PowerPoint.MsoAnimEffect.msoAnimEffectAscend},
        //    {"Blinds", PowerPoint.MsoAnimEffect.msoAnimEffectBlinds},
        //    {"Checkerboard", PowerPoint.MsoAnimEffect.msoAnimEffectCheckerboard},
        //    {"Circle", PowerPoint.MsoAnimEffect.msoAnimEffectCircle},
        //    {"Crawl", PowerPoint.MsoAnimEffect.msoAnimEffectCrawl},
        //    {"Credits", PowerPoint.MsoAnimEffect.msoAnimEffectCredits},
        //    {"Descend", PowerPoint.MsoAnimEffect.msoAnimEffectDescend},
        //    {"Diamond", PowerPoint.MsoAnimEffect.msoAnimEffectDiamond},
        //    {"Dissolve", PowerPoint.MsoAnimEffect.msoAnimEffectDissolve},
        //    {"Ease In", PowerPoint.MsoAnimEffect.msoAnimEffectEaseIn},
        //    {"Expand", PowerPoint.MsoAnimEffect.msoAnimEffectExpand},
        //    {"Fade", PowerPoint.MsoAnimEffect.msoAnimEffectFade},
        //    {"Faded Swivel", PowerPoint.MsoAnimEffect.msoAnimEffectFadedSwivel},
        //    {"Faded Zoom", PowerPoint.MsoAnimEffect.msoAnimEffectFadedZoom},
        //    {"Flash Bulb", PowerPoint.MsoAnimEffect.msoAnimEffectFlashBulb},
        //    {"Flash Once", PowerPoint.MsoAnimEffect.msoAnimEffectFlashOnce},
        //    {"Flicker", PowerPoint.MsoAnimEffect.msoAnimEffectFlicker},
        //    {"Flip", PowerPoint.MsoAnimEffect.msoAnimEffectFlip},
        //    {"Float", PowerPoint.MsoAnimEffect.msoAnimEffectFloat}
        //};

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
            if (tempSlide.Name.Contains("PPSlide") && tempSlide.Name.Substring(0, 7).Equals("PPSlide"))
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Slide nextSlide = presentation.Slides[tempSlide.SlideIndex + 1];
                PowerPoint.Slide currentSlide = presentation.Slides[tempSlide.SlideIndex - 1];
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.SlideIndex);
                tempSlide.Delete();

                AddCompleteAutoMotion(currentSlide, nextSlide);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("The current slide was not added by PPTLabs AutoMotion", "Error");
            }
        }

        public void AboutButtonClick(Office.IRibbonControl control)
        {
            AboutForm form = new AboutForm();
            form.Show();
        }

        public void HelpButtonClick(Office.IRibbonControl control)
        {
        }

        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
        }

        public void HighlightBulletsButtonClick(Office.IRibbonControl control)
        {
        }

        public void AddZoomButtonClick(Office.IRibbonControl control)
        {
        }

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
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
            if (tempSlide.Name.Contains("PPTLabsSpotlight")) //&& tempSlide.Name.Substring(0, 14).Equals("PPTLabsSpotlight")
            {
                PowerPoint.Slide currentSlide;
                PowerPoint.Shape spotlightShape;
                if (spotlightSlideMapping.TryGetValue(tempSlide.Name, out currentSlide))
                {
                    if (spotlightShapeMapping.TryGetValue(tempSlide.Name, out spotlightShape))
                    {
                        spotlightSlideMapping.Remove(tempSlide.Name);
                        spotlightShapeMapping.Remove(tempSlide.Name);
                        tempSlide.Delete();
                        AddSpotlightEffect(currentSlide, spotlightShape);
                    }
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
            PowerPoint.Shape spotlightShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            AddSpotlightEffect(currentSlide, spotlightShape);
            AddAckSlide();
        }

        #endregion

        #region Helpers

        private void AddSpotlightEffect(PowerPoint.Slide currentSlide, PowerPoint.Shape spotlightShape)
        {
            currentSlide.Duplicate();
            PowerPoint.Slide addedSlide = GetNextSlide(currentSlide);
            addedSlide.Name = "PPTLabsSpotlight" + GetTimestamp(DateTime.Now);
            spotlightSlideMapping.Add(addedSlide.Name, currentSlide);

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Shape rectangleShape = addedSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
            rectangleShape.Fill.ForeColor.RGB = 0x000000;
            rectangleShape.Fill.Transparency = defaultTransparency;
            rectangleShape.Line.Visible = Office.MsoTriState.msoFalse;
            rectangleShape.Name = "SpotlightShape1";

            PowerPoint.Shape selectedShape = spotlightShape;
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
            int color = newShape.Fill.ForeColor.RGB;
            //newShape.
            newShape.Fill.ForeColor.RGB = 0xffffff;
            newShape.Line.Visible = Office.MsoTriState.msoFalse;
            newShape.Name = "SpotlightShape2";
            spotlightShapeMapping.Add(addedSlide.Name, newShape);
            selectedShape.Delete();

            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            String[] array = { "SpotlightShape1", "SpotlightShape2" };
            PowerPoint.ShapeRange newRange = addedSlide.Shapes.Range(array);
            newRange.Select();

            PowerPoint.Selection currentSelection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            int count = currentSelection.ShapeRange.Count;
            currentSelection.Cut();

            PowerPoint.Shape pictureShape = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pictureShape.PictureFormat.TransparencyColor = 0xffffff;
            pictureShape.PictureFormat.TransparentBackground = Office.MsoTriState.msoTrue;
            pictureShape.SoftEdge.Radius = defaultSoftEdges;
            pictureShape.Width = presentation.PageSetup.SlideWidth + 2 * defaultSoftEdges;
            pictureShape.Height = presentation.PageSetup.SlideHeight + 2 * defaultSoftEdges;
            pictureShape.Left = 0 - defaultSoftEdges;
            pictureShape.Top = 0 - defaultSoftEdges;
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
            PowerPoint.Shape indicatorShape = newSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 200, 0, 200, 140); 
            //PowerPoint.Shape indicatorShape = newSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, presentation.PageSetup.SlideWidth - 100, 0, 100.0f, 100.0f);
            //indicatorShape.ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStylePreset28;
            //indicatorShape.TextFrame2.TextRange.Text = "Added By \nPowerPointLabs";
            //indicatorShape.TextFrame2.NoTextRotation = Office.MsoTriState.msoFalse;
            //indicatorShape.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
            //indicatorShape.TextFrame2.WordWrap = Office.MsoTriState.msoFalse;

            //indicatorShape.Rotation = 180;
            //indicatorShape.Cut();
            //indicatorShape = newSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            //indicatorShape.Left = presentation.PageSetup.SlideWidth - indicatorShape.Width + 5;
            //indicatorShape.Top = 0;
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

                        //Motion Effect
                        float finalX = (shapes2[count].Left + (shapes2[count].Width) / 2);
                        float initialX = (shapes1[count].Left + (shapes1[count].Width) / 2);
                        float finalY = (shapes2[count].Top + (shapes2[count].Height) / 2);
                        float initialY = (shapes1[count].Top + (shapes1[count].Height) / 2);

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
                            motion.MotionEffect.Path = "M 0 0 C " + ((finalX - initialX) / 2) / presentation.PageSetup.SlideWidth + " " + ((finalY - initialY) / 2) / presentation.PageSetup.SlideHeight + " " + ((finalX - initialX) / 2) / presentation.PageSetup.SlideWidth + " " + ((finalY - initialY) / 2) / presentation.PageSetup.SlideHeight + " " + (finalX - initialX) / presentation.PageSetup.SlideWidth + " " + (finalY - initialY) / presentation.PageSetup.SlideHeight + " E";
                        }

                        //Resize Effect
                        if (sh.Type != Office.MsoShapeType.msoPlaceholder && sh.Type != Office.MsoShapeType.msoTextBox)
                        {
                            float finalWidth = shapes2[count].Width;
                            float initialWidth = shapes1[count].Width;
                            float finalHeight = shapes2[count].Height;
                            float initialHeight = shapes1[count].Height;

                            if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
                            {
                                effectResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                                //PowerPoint.AnimationBehavior resize = effectResize.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeScale);
                                PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];

                                effectResize.Timing.Duration = defaultDuration;
                                resize.ScaleEffect.ToX = (finalWidth / initialWidth) * 100;
                                resize.ScaleEffect.ToY = (finalHeight / initialHeight) * 100;
                                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                            }
                        }
                        if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame.HasText == Office.MsoTriState.msoTrue && sh.TextFrame.TextRange.Font.Size != shapes2[count].TextFrame.TextRange.Font.Size)
                        {
                            sh.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
                            effectFontResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectChangeFontSize, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            effectFontResize.Timing.Duration = defaultDuration;
                            PowerPoint.AnimationBehavior resizeFont = effectFontResize.Behaviors[1];
                            resizeFont.PropertyEffect.To = shapes2[count].TextFrame.TextRange.Font.Size / shapes1[count].TextFrame.TextRange.Font.Size;
                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        }

                        //Rotation Effect
                        float finalRotation = shapes2[count].Rotation;
                        float initialRotation = shapes1[count].Rotation;
                        if (finalRotation != initialRotation)
                        {
                            effectRotate = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            effectRotate.Timing.Duration = defaultDuration;
                            effectRotate.EffectParameters.Amount = GetMinimumRotation(initialRotation, finalRotation);
                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
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
            //Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, currentSlide.CustomLayout);

            //Store reference of new slide
            PowerPoint.Slide newSlide = GetNextSlide(currentSlide);
            newSlide.Name = "PPSlide" + GetTimestamp(DateTime.Now);

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

            AddAckSlide();

            return newSlide;
        }

        private void AddAckSlide()
        {
            PowerPoint.Slide tempSlide = (Globals.ThisAddIn.Application.ActivePresentation.Slides[Globals.ThisAddIn.Application.ActivePresentation.Slides.Count]);
            if (!(tempSlide.Name.Contains("PPAck") && tempSlide.Name.Substring(0, 5).Equals("PPAck")))
            {
                PowerPoint.Slide ackSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(Globals.ThisAddIn.Application.ActivePresentation.Slides.Count + 1, GetCurrentSlide().CustomLayout);
                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(ackSlide.SlideIndex);
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
                        if (sh1.Type == Office.MsoShapeType.msoPlaceholder && sh1.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            sh1.TextFrame.TextRange.Text.Trim();
                            sh1.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                        if (sh2.Type == Office.MsoShapeType.msoPlaceholder && sh2.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            sh2.TextFrame.TextRange.Text.Trim();
                            sh2.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }

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
            float rotation2 = rotation1 == 0 ? 0 : Math.Abs(360 - Math.Abs(rotation1)) * (rotation1 / Math.Abs(rotation1)) * -1;

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
            float d = Math.Abs(i) % 360;

            if (i < 0)
            {
                //return positive equivalent
                return 360 - d;
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
