using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
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
        public bool frameAnimationChecked = false;
        public bool backgroundZoomChecked = false;
        public bool spotlightDelete = true;
        public float defaultSoftEdges = 10;
        public float defaultDuration = 0.5f;
        public float defaultTransparency = 0.7f;
        public bool startUp = false;
        public bool spotlightEnabled = false;
        public bool inSlideEnabled = false;
        public bool zoomButtonEnabled = false;
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

        private bool _allSlides;
        private bool _previewCurrentSlide;
        private bool _captionsAllSlides;

        private IEnumerable<string> _voiceNames;

        private int _voiceSelected = 0;

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

            LoadVoicesIntoDropdown();
            SetCoreVoicesToSelections();
        }
        public void RefreshRibbonControl(String controlID)
        {
            try
            {
                ribbon.InvalidateControl(controlID);
            }
            catch (Exception e)
            {
                LogException(e, "RefreshRibbonControl");
                throw;
            }
        }

        private void LoadVoicesIntoDropdown()
        {
            SetVoicesFromInstalledOptions();
            RefreshVoicePicker();
        }

        private void SetVoicesFromInstalledOptions()
        {
            var installedVoices = NotesToAudio.GetVoices().ToList();
            _voiceNames = installedVoices;
        }

        private void RefreshVoicePicker()
        {
            RefreshRibbonControl("defaultVoicePicker");
        }

        //Button Click Callbacks
        private PowerPoint.Shape FindIdenticalShape(PowerPoint.Slide slideToSearch, PowerPoint.Shape shapeToSearch)
        {
            PowerPoint.Shape shapeToReturn = null;
            foreach (PowerPoint.Shape sh in slideToSearch.Shapes)
            {
                if (sh.Id == shapeToSearch.Id && sh.Name.Equals(shapeToSearch.Name))
                {
                    shapeToReturn = sh;
                    break;
                }
            }
            return shapeToReturn;
        }
        private PowerPoint.Effect GetShapeAnnimations(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            try
            {
                PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
                PowerPoint.Effect e = null;
                for (int x = sequence.Count; x >= 1; x--)
                {
                    PowerPoint.Effect effect = sequence[x];
                    if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    {
                        e = effect;
                        break;
                    }
                }
                return e;
            }
            catch (Exception e)
            {
                LogException(e, "GetShapeAnimations");
                throw;
            }
        }
        public void AddInSlideAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                //Get References of current and next slides
                PowerPoint.Slide currentSlide = GetCurrentSlide();
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count == 0)
                    return;

                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.ShapeRange shapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;

                foreach (PowerPoint.Shape sh in shapes)
                {
                    DeleteShapeAnnimations(currentSlide, sh);
                }

                PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
                PowerPoint.Effect effectMotion = null;
                PowerPoint.Effect effectResize = null;
                PowerPoint.Effect effectRotate = null;
                PowerPoint.MsoAnimTriggerType trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;

                for (int num = 1; num <= shapes.Count - 1; num++)
                {
                    PowerPoint.Shape sh1 = shapes[num];
                    PowerPoint.Shape sh2 = shapes[num + 1];

                    if (sh1 == null || sh2 == null)
                        return;

                    if (num == 1)
                    {
                        PowerPoint.Effect appear = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    }

                    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                    float finalX = (sh2.Left + (sh2.Width) / 2);
                    float initialX = (sh1.Left + (sh1.Width) / 2);
                    float finalY = (sh2.Top + (sh2.Height) / 2);
                    float initialY = (sh1.Top + (sh1.Height) / 2);

                    float finalRotation = sh2.Rotation;
                    float initialRotation = sh1.Rotation;

                    float finalWidth = sh2.Width;
                    float initialWidth = sh1.Width;
                    float finalHeight = sh2.Height;
                    float initialHeight = sh1.Height;
                    float finalFont = 0.0f;
                    float initialFont = 0.0f;
                    int numFrames = (int)(defaultDuration / 0.04f);
                    numFrames = (numFrames > 30) ? 30 : numFrames;

                    if (sh1.HasTextFrame == Office.MsoTriState.msoTrue && (sh1.TextFrame.HasText == Office.MsoTriState.msoTriStateMixed || sh1.TextFrame.HasText == Office.MsoTriState.msoTrue) && sh1.TextFrame.TextRange.Font.Size != sh2.TextFrame.TextRange.Font.Size)
                    {
                        finalFont = sh2.TextFrame.TextRange.Font.Size;
                        initialFont = sh1.TextFrame.TextRange.Font.Size;
                    }

                    if ((frameAnimationChecked && (finalHeight != initialHeight || finalWidth != initialWidth))
                        || ((initialRotation != finalRotation || initialRotation % 90 != 0) && (finalHeight != initialHeight || finalWidth != initialWidth))
                        || finalFont != initialFont)
                    {
                        float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
                        float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
                        float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
                        float incrementLeft = (finalX - initialX) / numFrames;
                        float incrementTop = (finalY - initialY) / numFrames;
                        float incrementFont = (finalFont - initialFont) / numFrames;

                        //PowerPoint.Effect shapeEffect = GetShapeAnnimations(addedSlide, sh1);
                        //if (shapeEffect != null)
                        //    shapeEffect.Delete();

                        PowerPoint.Shape lastShape = sh1;
                        for (int i = 1; i <= numFrames; i++)
                        {
                            PowerPoint.Shape dupShape = sh1.Duplicate()[1];
                            if (i != 1)
                            {
                                sequence[sequence.Count].Delete();
                            }
                            PowerPoint.Effect shapeEffect = GetShapeAnnimations(currentSlide, dupShape);
                            if (shapeEffect != null)
                                shapeEffect.Delete();

                            dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                            dupShape.Left = sh1.Left;
                            dupShape.Top = sh1.Top;

                            if (incrementWidth != 0.0f)
                            {
                                dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                            }

                            if (incrementHeight != 0.0f)
                            {
                                dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                            }

                            if (incrementRotation != 0.0f)
                            {
                                dupShape.Rotation += (incrementRotation * i);
                            }

                            if (incrementLeft != 0.0f)
                            {
                                dupShape.Left += (incrementLeft * i);
                            }

                            if (incrementTop != 0.0f)
                            {
                                dupShape.Top += (incrementTop * i);
                            }

                            if (incrementFont != 0.0f)
                            {
                                dupShape.TextFrame.TextRange.Font.Size += (incrementFont * i);
                            }

                            if (i == 1)
                            {
                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                            }
                            else
                            {
                                PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                appear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);
                            }

                            PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            disappear.Exit = Office.MsoTriState.msoTrue;
                            disappear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                            lastShape = dupShape;
                        }
                        PowerPoint.Effect disappearLast = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        disappearLast.Exit = Office.MsoTriState.msoTrue;
                        disappearLast.Timing.TriggerDelayTime = defaultDuration;
                    }
                    else
                    {
                        //Motion Effect
                        if ((finalX != initialX) || (finalY != initialY))
                        {
                            effectMotion = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                            effectMotion.Timing.Duration = defaultDuration;
                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;

                            //Create VML path for the motion path
                            //This path needs to be a curved path to allow the user to edit points
                            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;
                            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;
                        }

                        //Resize Effect
                        if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
                        {
                            sh1.LockAspectRatio = Office.MsoTriState.msoFalse;
                            effectResize = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                            effectResize.Timing.Duration = defaultDuration;

                            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        }

                        //Rotation Effect
                        if (finalRotation != initialRotation)
                        {
                            effectRotate = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                            PowerPoint.AnimationBehavior rotate = effectRotate.Behaviors[1];
                            effectRotate.Timing.Duration = defaultDuration;
                            effectRotate.EffectParameters.Amount = GetMinimumRotation(initialRotation, finalRotation);
                            trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        }
                    }

                    PowerPoint.Effect shape2Appear = sequence.AddEffect(sh2, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);

                    PowerPoint.Effect shape1Disappear = sequence.AddEffect(sh1, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    shape1Disappear.Exit = Office.MsoTriState.msoTrue;
                }
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                AddAckSlide();
            }
            catch (Exception e)
            {
                LogException(e, "AddInSlideAnimationButtonClick");
                throw;
            }
        }
        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            try
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
            catch (Exception e)
            {
                LogException(e, "AddAnimationButtonClick");
                throw;
            }
        }
        public void ReloadButtonClick(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide tempSlide = GetCurrentSlide();
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                if (tempSlide.Name.Contains("PPSlideAnimated") && tempSlide.Name.Substring(0, 15).Equals("PPSlideAnimated"))
                {
                    PowerPoint.Slide nextSlide = presentation.Slides[tempSlide.SlideIndex + 1];
                    PowerPoint.Slide currentSlide = presentation.Slides[tempSlide.SlideIndex - 1];
                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.SlideIndex);
                    tempSlide.Delete();

                    AddCompleteAutoMotion(currentSlide, nextSlide);
                }
                else if (tempSlide.Name.Contains("PPSlideStart") && tempSlide.Name.Substring(0, 12).Equals("PPSlideStart"))
                {
                    PowerPoint.Slide animatedSlide = presentation.Slides[tempSlide.SlideIndex + 1];
                    PowerPoint.Slide finalSlide = presentation.Slides[tempSlide.SlideIndex + 2];
                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);
                    animatedSlide.Delete();

                    AddCompleteAutoMotion(tempSlide, finalSlide);
                }
                else if (tempSlide.Name.Contains("PPSlideEnd") && tempSlide.Name.Substring(0, 10).Equals("PPSlideEnd"))
                {
                    PowerPoint.Slide animatedSlide = presentation.Slides[tempSlide.SlideIndex - 1];
                    PowerPoint.Slide firstSlide = presentation.Slides[tempSlide.SlideIndex - 2];
                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(tempSlide.SlideIndex);
                    animatedSlide.Delete();

                    AddCompleteAutoMotion(firstSlide, tempSlide);
                }
                else if (tempSlide.Name.Contains("PPSlideMulti") && tempSlide.Name.Substring(0, 12).Equals("PPSlideMulti"))
                {
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
            catch (Exception e)
            {
                LogException(e, "ReloadAnimationButtonClick");
                throw;
            }
        }
        private PowerPoint.Slide AddMagnifyingSlide(PowerPoint.Slide currentSlide, PowerPoint.Shape selectedShape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = currentSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsMagnifyingSlide" + GetTimestamp(DateTime.Now);

            float centerX = selectedShape.Left + selectedShape.Width / 2;
            float centerY = selectedShape.Top + selectedShape.Height / 2;

            PowerPoint.Shape sh = FindIdenticalShape(addedSlide, selectedShape);
            if (sh != null)
            {
                sh.Delete();
            }

            addedSlide.Copy();
            PowerPoint.Shape duplicatePic = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            duplicatePic.Name = "PPTLabsMagnifyArea" + GetTimestamp(DateTime.Now);

            float scaleFactorX = presentation.PageSetup.SlideWidth / duplicatePic.Width;
            float scaleFactorY = presentation.PageSetup.SlideHeight / duplicatePic.Height;

            duplicatePic.LockAspectRatio = Office.MsoTriState.msoFalse;
            duplicatePic.Left = 0;
            duplicatePic.Top = 0;
            duplicatePic.Width = presentation.PageSetup.SlideWidth;
            duplicatePic.Height = presentation.PageSetup.SlideHeight;

            duplicatePic.PictureFormat.CropLeft += selectedShape.Left / scaleFactorX;
            duplicatePic.PictureFormat.CropTop += selectedShape.Top / scaleFactorY;
            duplicatePic.PictureFormat.CropRight += (presentation.PageSetup.SlideWidth - (selectedShape.Left + selectedShape.Width)) / scaleFactorX;
            duplicatePic.PictureFormat.CropBottom += (presentation.PageSetup.SlideHeight - (selectedShape.Top + selectedShape.Height)) / scaleFactorY;

            //duplicatePic.Cut();

            //currentSlide.Duplicate();
            //Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, Globals.ThisAddIn.Application.ActivePresentation.SlideMaster.CustomLayouts[7]);
            //PowerPoint.Slide addedSlide = GetNextSlide(currentSlide);

            //Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            //PowerPoint.Shape sh = addedSlide.Shapes.Paste()[1];
            duplicatePic.Left = centerX - (duplicatePic.Width / 2);
            duplicatePic.Top = centerY - (duplicatePic.Height / 2);

            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectFade = null;

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;

            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(duplicatePic) || tmp.Equals(indicatorShape)))
                {
                    DeleteShapeAnnimations(addedSlide, tmp);
                    effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.5f;
                }
            }

            float finalX = (presentation.PageSetup.SlideWidth / 2);
            float initialX = (duplicatePic.Left + (duplicatePic.Width) / 2);
            float finalY = (presentation.PageSetup.SlideHeight / 2);
            float initialY = (duplicatePic.Top + (duplicatePic.Height) / 2);

            float finalWidth = presentation.PageSetup.SlideWidth;
            float initialWidth = duplicatePic.Width;
            float finalHeight = presentation.PageSetup.SlideHeight;
            float initialHeight = duplicatePic.Height;

            effectMotion = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
            effectMotion.Timing.Duration = 0.5f;
            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

            effectResize = sequence.AddEffect(duplicatePic, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
            effectResize.Timing.Duration = 0.5f;
            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceTime = 0;
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Shape tmp in addedSlide.NotesPage.Shapes)
                {
                    if (tmp.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        tmp.TextEffect.Text = "";
                }
            }

            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (tmp.Type == Office.MsoShapeType.msoMedia)
                    tmp.Delete();
            }
            return addedSlide;
        }
        private PowerPoint.Slide AddMagnifiedSlide(PowerPoint.Slide magnifyingSlide)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = magnifyingSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsMagnifiedSlide" + GetTimestamp(DateTime.Now);
            PowerPoint.Shape magnifyShape = null;
            foreach (PowerPoint.Shape sh in magnifyingSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                if (!sh.Name.Contains("PPTLabsMagnifyArea"))
                {
                    tmp.Delete();
                }
                else
                {
                    magnifyShape = tmp;
                }
            }

            DeleteShapeAnnimations(addedSlide, magnifyShape);
            magnifyShape.LockAspectRatio = Office.MsoTriState.msoFalse;
            magnifyShape.Left = 0;
            magnifyShape.Top = 0;
            magnifyShape.Width = presentation.PageSetup.SlideWidth;
            magnifyShape.Height = presentation.PageSetup.SlideHeight;

            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
            return addedSlide;
        }
        private PowerPoint.Slide AddDeMagnifyingSlide(PowerPoint.Slide magnifyingSlide, PowerPoint.Shape selectedShape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide addedSlide = magnifyingSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsDeMagnifyingSlide" + GetTimestamp(DateTime.Now);
            addedSlide.MoveTo(magnifyingSlide.SlideIndex + 2);

            PowerPoint.Shape magnifyShape = null;
            foreach (PowerPoint.Shape sh in magnifyingSlide.Shapes)
            {
                PowerPoint.Shape tmp = FindIdenticalShape(addedSlide, sh);
                DeleteShapeAnnimations(addedSlide, tmp);
                if (sh.Name.Contains("PPTLabsMagnifyArea"))
                {
                    magnifyShape = tmp;
                }
                if (sh.Name.Contains("PPIndicator"))
                {
                    tmp.Delete();
                }
            }

            magnifyShape.LockAspectRatio = Office.MsoTriState.msoFalse;
            magnifyShape.Left = 0;
            magnifyShape.Top = 0;
            magnifyShape.Width = presentation.PageSetup.SlideWidth;
            magnifyShape.Height = presentation.PageSetup.SlideHeight;

            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectFade = null;

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            String tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            float finalX = (selectedShape.Left + (selectedShape.Width) / 2);
            float initialX = (magnifyShape.Left + (magnifyShape.Width) / 2);
            float finalY = (selectedShape.Top + (selectedShape.Height) / 2);
            float initialY = (magnifyShape.Top + (magnifyShape.Height) / 2);

            float finalWidth = selectedShape.Width;
            float initialWidth = magnifyShape.Width;
            float finalHeight = selectedShape.Height;
            float initialHeight = magnifyShape.Height;

            effectMotion = sequence.AddEffect(magnifyShape, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
            effectMotion.Timing.Duration = 0.5f;
            float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
            float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
            float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
            float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

            motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
            effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
            effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

            effectResize = sequence.AddEffect(magnifyShape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
            effectResize.Timing.Duration = 0.5f;
            resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
            resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

            int i = 0;
            foreach (PowerPoint.Shape tmp in addedSlide.Shapes)
            {
                if (!(tmp.Equals(magnifyShape) || tmp.Equals(indicatorShape)))
                {
                    if (i == 0)
                    {
                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                    }
                    else
                    {
                        effectFade = sequence.AddEffect(tmp, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    }
                    effectFade.Timing.Duration = 0.5f;
                    i++;
                }
            }
            effectFade = sequence.AddEffect(magnifyShape, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectFade.Exit = Office.MsoTriState.msoTrue;
            effectFade.Timing.Duration = 0.5f;

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;

            return addedSlide;
        }
        public void ZoomBtnClick(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Slide currentSlide = GetCurrentSlide();
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count == 0)
                    return;

                PowerPoint.Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = "PPTLabsMagnifyShape" + GetTimestamp(DateTime.Now);

                PowerPoint.Slide magnifyingSlide = AddMagnifyingSlide(currentSlide, selectedShape);
                PowerPoint.Slide magnifiedSlide = AddMagnifiedSlide(magnifyingSlide);
                PowerPoint.Slide demagnifyingSlide = AddDeMagnifyingSlide(magnifyingSlide, selectedShape);
                selectedShape.Visible = Office.MsoTriState.msoFalse;

                Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(magnifyingSlide.SlideIndex);
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                AddAckSlide();
            }
            catch (Exception e)
            {
                LogException(e, "ZoomBtnClick");
                throw;
            }
        }
        public void ReloadSpotlightButtonClick(Office.IRibbonControl control)
        {
            try
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
            catch (Exception e)
            {
                LogException(e, "ReloadSpotlightButtonClick");
                throw;
            }
        }
        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            try
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
                            //if (spotlightDelete)
                            //{
                                copyShape.Delete();
                            //}
                            //else
                            //{
                            //    copyShape.Name = "SpotlightCopy" + GetTimestamp(DateTime.Now);
                            //}

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

                    //if (!spotlightDelete)
                    //{
                    //    float oldLeft = spotlightShape.Left;
                    //    float oldTop = spotlightShape.Top;
                    //    spotlightShape.Left -= 5.0f;
                    //    spotlightShape.Top -= 5.0f;
                    //    spotlightShape.Width += 10.0f;
                    //    spotlightShape.Height += 10.0f;

                    //    if (spotlightShape.Left < 0.0f)
                    //    {
                    //        spotlightShape.Left = 0.0f;
                    //        spotlightShape.Width = spotlightShape.Width - (0.0f - spotlightShape.Left);
                    //    }
                    //    if (spotlightShape.Top < 0.0f)
                    //    {
                    //        spotlightShape.Top = 0.0f;
                    //        spotlightShape.Top = spotlightShape.Height - (0.0f - spotlightShape.Top);
                    //    }
                    //    if (spotlightShape.Left + spotlightShape.Width > Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth)
                    //        spotlightShape.Width = (Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth - oldLeft);
                    //    if (spotlightShape.Top + spotlightShape.Height > Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight)
                    //        spotlightShape.Height = (Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight - oldTop);
                    //}
                    
                    spotlightShape.Fill.ForeColor.RGB = 0xffffff;
                    spotlightShape.Line.Visible = Office.MsoTriState.msoFalse;
                    if (spotlightShape.HasTextFrame == Office.MsoTriState.msoTrue && spotlightShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        spotlightShape.TextFrame.TextRange.Font.Color.RGB = 0xffffff;
                    spotlightShape.Name = "SpotlightShape" + counter;
                    counter++;

                    PowerPoint.Shape duplicateShape = spotlightShape.Duplicate()[1];
                    duplicateShape.Visible = Office.MsoTriState.msoFalse;
                    duplicateShape.Left = spotlightShape.Left;
                    duplicateShape.Top = spotlightShape.Top;

                    spotlightShapes.Add(spotlightShape);
                    //if (spotlightDelete)
                    spotShape.Delete();
                }

                if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
                {
                    foreach (PowerPoint.Shape sh in addedSlide.NotesPage.Shapes)
                    {
                        if (sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                            sh.TextEffect.Text = "";
                    }
                }

                foreach (PowerPoint.Shape sh in addedSlide.Shapes)
                {
                    if (sh.Type == Office.MsoShapeType.msoMedia)
                        sh.Delete();
                }

                AddSpotlightEffect(addedSlide, spotlightShapes);
                AddAckSlide();

                //Bring spotlight shapes to front
                //List<PowerPoint.Shape> shapesToEdit = new List<PowerPoint.Shape>();
                //foreach (PowerPoint.Shape copyShape in addedSlide.Shapes)
                //{
                //    if (copyShape.Name.Contains("SpotlightCopy") && !spotlightDelete)
                //    {
                //        copyShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                //    }
                //}
            }
            catch (Exception e)
            {
                LogException(e, "SpotlightBtnClick");
                throw;
            }
        }
        public void AboutButtonClick(Office.IRibbonControl control)
        {
            //AboutForm form = new AboutForm();
            //form.Show();
            System.Windows.Forms.MessageBox.Show("          PowerPointLabs Plugin Version 1.4.3 [Release date: 24 Feb 2014]\n     Developed at School of Computing, National University of Singapore.\n        For more information, visit our website http://PowerPointLabs.info", "About PowerPointLabs");
        }
        public void HelpButtonClick(Office.IRibbonControl control)
        {
            try
            {
                string myURL = "http://powerpointlabs.info/docs.html";
                System.Diagnostics.Process.Start(myURL);
            }
            catch (Exception e)
            {
                LogException(e, "HelpButtonClick");
                throw;
            }
        }
        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
            try
            {
                string myURL = "http://powerpointlabs.info/contact.html";
                System.Diagnostics.Process.Start(myURL);
            }
            catch (Exception e)
            {
                LogException(e, "FeedbackButtonClick");
                throw;
            }
        }
        public void HighlightBulletsButtonClick(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("This feature is coming soon!                ", "Coming Soon");
        }
        private PowerPoint.Shape FindMatchingShape(PowerPoint.Slide slideToSearch, PowerPoint.Shape shapeToSearch)
        {
            PowerPoint.Shape shapeToReturn = null;
            foreach (PowerPoint.Shape sh in slideToSearch.Shapes)
            {
                if (sh.Name.Equals(shapeToSearch.Name))
                {
                    shapeToReturn = sh;
                    break;
                }
            }
            return shapeToReturn;
        }
        private void ZoomInOmitBackground(PowerPoint.Slide currentSlide, PowerPoint.Shape shape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide nextSlide = GetNextSlide(currentSlide);
            PowerPoint.Slide tempSlide = nextSlide;
            if (nextSlide.Name.Contains("PPTLabsZoomIn") && nextSlide.SlideIndex < presentation.Slides.Count)
            {
                nextSlide = GetNextSlide(tempSlide);
                tempSlide.Delete();
            }

            if (nextSlide.Shapes.Count == 0)
                return;

            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.SlideIndex);
            String tempFileName = "";
            foreach (PowerPoint.Shape sh in nextSlide.Shapes)
            {
                sh.Copy();
                tempFileName = Path.GetTempFileName();
                PowerPoint.Shape tempPicture1 = nextSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                tempPicture1.Export(tempFileName, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                tempPicture1.Delete();

                PowerPoint.Shape tempPicture2 = currentSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, sh.Left, sh.Top, sh.Width, sh.Height);
                sh.Name += GetTimestamp(DateTime.Now);
                tempPicture2.Name = sh.Name;
                tempPicture2.Select(Office.MsoTriState.msoFalse);
            }
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape shapeGroup = null;
            if (sel.ShapeRange.Count > 1)
            {
                shapeGroup = sel.ShapeRange.Group();
            }
            else
            {
                shapeGroup = sel.ShapeRange[1];
            }

            shapeGroup.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeGroup.Left = shape.Left;
            shapeGroup.Top = shape.Top;
            shapeGroup.Width = shape.Width;
            shapeGroup.Height = shape.Height;
            //shapeGroup.Name = "PPTZoomInShape" + GetTimestamp(DateTime.Now);
            shape.Visible = Office.MsoTriState.msoFalse;

            PowerPoint.Slide addedSlide = currentSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsZoomIn" + GetTimestamp(DateTime.Now);

            //PowerPoint.ShapeRange range = null;
            PowerPoint.Shape identicalShape = FindIdenticalShape(addedSlide, shapeGroup);
            List<PowerPoint.Shape> shapesToZoom = new List<PowerPoint.Shape>();
            if (identicalShape.Type == Office.MsoShapeType.msoGroup)
            {
                //range = identicalShape.Ungroup();
                foreach (PowerPoint.Shape tmp in identicalShape.GroupItems)
                {
                    shapesToZoom.Add(tmp);
                }
            }
            else
            {
                shapesToZoom.Add(identicalShape);
            }

            shapeGroup.Cut();
            PowerPoint.Shape shapePicture = currentSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            shapePicture.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapePicture.Left = shape.Left;
            shapePicture.Top = shape.Top;
            shapePicture.Width = shape.Width;
            shapePicture.Height = shape.Height;

            PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
            PowerPoint.Effect effectAppear = sequence.AddEffect(shapePicture, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            effectAppear.Timing.Duration = 0.50f;
            sequence = addedSlide.TimeLine.MainSequence;

            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectFade = null;

            tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;
            bool fadeFlag = false;
            int i = 0;

            foreach (PowerPoint.Shape sh in addedSlide.Shapes)
            {
                if (!shapesToZoom.Contains(sh) && !sh.Equals(indicatorShape) && !sh.Equals(identicalShape))
                {
                    effectFade = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectFade.Exit = Office.MsoTriState.msoTrue;
                    effectFade.Timing.Duration = 0.5f;
                    fadeFlag = true;
                }
            }

            foreach (PowerPoint.Shape sh in shapesToZoom)
            {
                PowerPoint.MsoAnimTriggerType trigger = (i == 0 && fadeFlag) ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious : PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                PowerPoint.Shape matchingShape = FindMatchingShape(nextSlide, sh);
                matchingShape.Name = matchingShape.Name.Substring(0, matchingShape.Name.Length - 18);
                float finalWidth = matchingShape.Width;
                float initialWidth = sh.Width;
                float finalHeight = matchingShape.Height;
                float initialHeight = sh.Height;

                float finalX = (matchingShape.Left + (matchingShape.Width) / 2);
                float initialX = (sh.Left + (sh.Width) / 2);
                float finalY = (matchingShape.Top + (matchingShape.Height) / 2);
                float initialY = (sh.Top + (sh.Height) / 2);

                effectMotion = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                effectMotion.Timing.Duration = defaultDuration;
                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

                motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

                effectResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                effectResize.Timing.Duration = defaultDuration;
                resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;
                trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;

                i++;
            }            

            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceTime = 0;
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Shape sh in addedSlide.NotesPage.Shapes)
                {
                    if (sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        sh.TextEffect.Text = "";
                }
            }

            foreach (PowerPoint.Shape sh in addedSlide.Shapes)
            {
                if (sh.Type == Office.MsoShapeType.msoMedia)
                    sh.Delete();
            }

            nextSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
            nextSlide.SlideShowTransition.Duration = 0.25f;
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
            AddAckSlide();
        }
        public void AddZoomInButtonClick(Office.IRibbonControl control)
        {
            //System.Windows.Forms.MessageBox.Show("This feature is coming soon!                ", "Coming Soon");
            
            //Get References of current and next slides
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count == 0)
                return;

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Shape shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            DeleteShapeAnnimations(currentSlide, shape);
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
            {
                if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Text = "";
                }
            }
            shape.Rotation = 0;

            if (currentSlide != null && currentSlide.SlideIndex != presentation.Slides.Count)
            {
                if (!backgroundZoomChecked)
                {
                    ZoomInOmitBackground(currentSlide, shape);
                }
                else
                {
                    PowerPoint.Slide nextSlide = GetNextSlide(currentSlide);
                    PowerPoint.Slide tempSlide = nextSlide;
                    if (nextSlide.Name.Contains("PPTLabsZoomIn") && nextSlide.SlideIndex < presentation.Slides.Count)
                    {
                        nextSlide = GetNextSlide(tempSlide);
                        tempSlide.Delete();
                    }

                    String tempFileName = Path.GetTempFileName();
                    nextSlide.Export(tempFileName, "PNG");
                    shape.Fill.UserPicture(tempFileName);
                    shape.Line.Visible = Office.MsoTriState.msoFalse;
                    shape.Name = "PPTZoomInShape" + GetTimestamp(DateTime.Now);

                    PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
                    PowerPoint.Effect effectAppear = sequence.AddEffect(shape, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    effectAppear.Timing.Duration = 0.50f;

                    PowerPoint.Slide addedSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, currentSlide.CustomLayout);
                    addedSlide.Name = "PPTLabsZoomIn" + GetTimestamp(DateTime.Now);

                    currentSlide.Copy();
                    //tempFileName = Path.GetTempFileName() + ".png";
                    //currentSlide.Export(tempFileName, "PNG");
                    //PowerPoint.Shape zoomShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0, 0);
                    PowerPoint.Shape zoomShape = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                    sequence = addedSlide.TimeLine.MainSequence;

                    zoomShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                    zoomShape.Left = 0;
                    zoomShape.Top = 0;
                    zoomShape.Width = presentation.PageSetup.SlideWidth;
                    zoomShape.Height = presentation.PageSetup.SlideHeight;
                    zoomShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    PowerPoint.Effect effectMotion = null;
                    PowerPoint.Effect effectResize = null;
                    //PowerPoint.Effect effectRotate = null;
                    PowerPoint.Effect effectDisappear = null;

                    tempFileName = Path.GetTempFileName();
                    Properties.Resources.Indicator.Save(tempFileName);
                    PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
                    indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
                    indicatorShape.Top = 0;
                    indicatorShape.Width = 120;
                    indicatorShape.Height = 84;
                    indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
                    effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectDisappear.Exit = Office.MsoTriState.msoTrue;
                    effectDisappear.Timing.Duration = 0;

                    //float angle = GetMinimumRotation(shape.Rotation, 0) * (float)Math.PI / 180.0f;

                    float finalWidth = presentation.PageSetup.SlideWidth;
                    float initialWidth = shape.Width;
                    float finalHeight = presentation.PageSetup.SlideHeight;
                    float initialHeight = shape.Height;

                    float finalX = (presentation.PageSetup.SlideWidth / 2) * (finalWidth / initialWidth);
                    float initialX = (shape.Left + (shape.Width) / 2) * (finalWidth / initialWidth);
                    float finalY = (presentation.PageSetup.SlideHeight / 2) * (finalHeight / initialHeight);
                    float initialY = (shape.Top + (shape.Height) / 2) * (finalHeight / initialHeight);

                    //float s = (float)Math.Sin(angle);
                    //float c = (float)Math.Cos(angle);

                    //finalX -= initialX;
                    //finalY -= initialY;

                    //// rotate point
                    //float newX = (finalX * c - finalY * s);
                    //float newY = (finalX * s + finalY * c);

                    //finalX = newX + initialX;
                    //finalY = newY + initialY;

                    effectMotion = sequence.AddEffect(zoomShape, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                    effectMotion.Timing.Duration = defaultDuration;
                    float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                    float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                    float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                    float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                    float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                    float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

                    motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                    effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                    effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

                    effectResize = sequence.AddEffect(zoomShape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                    effectResize.Timing.Duration = defaultDuration;
                    resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                    resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                    //effectRotate = sequence.AddEffect(zoomShape, PowerPoint.MsoAnimEffect.msoAnimEffectSpin, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    //PowerPoint.AnimationBehavior rotate = effectRotate.Behaviors[1];
                    //effectRotate.Timing.Duration = defaultDuration;
                    //effectRotate.EffectParameters.Amount = GetMinimumRotation(shape.Rotation, 0);

                    indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
                    addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
                    addedSlide.SlideShowTransition.AdvanceTime = 0;
                    addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
                    if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
                    {
                        foreach (PowerPoint.Shape sh in addedSlide.NotesPage.Shapes)
                        {
                            if (sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                sh.TextEffect.Text = "";
                        }
                    }

                    foreach (PowerPoint.Shape sh in addedSlide.Shapes)
                    {
                        if (sh.Type == Office.MsoShapeType.msoMedia)
                            sh.Delete();
                    }

                    //if (nextSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFade && nextSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFadeSmoothly)
                    //    nextSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
                    nextSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
                    nextSlide.SlideShowTransition.Duration = 0.25f;
                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                    AddAckSlide();
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select the correct slide", "Unable to Add Animations");
            }
        }
        private PowerPoint.Shape convertShapeToPicture(PowerPoint.Slide slide, PowerPoint.Shape sh)
        {
            sh.Copy();
            PowerPoint.Shape tempPicture1 = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            tempPicture1.LockAspectRatio = Office.MsoTriState.msoFalse;
            tempPicture1.Left = sh.Left;
            tempPicture1.Top = sh.Top;
            tempPicture1.Width = sh.Width;
            tempPicture1.Height = sh.Height;
            tempPicture1.Name = sh.Name;
            //sh.Delete();
            return tempPicture1;

        }
        private void ZoomOutOmitBackground(PowerPoint.Slide currentSlide, PowerPoint.Shape shape)
        {
            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide prevSlide = GetPrevSlide(currentSlide);
            PowerPoint.Slide tempSlide = prevSlide;
            while (prevSlide.Name.Contains("PPTLabsZoomOut") && prevSlide.SlideIndex > 1)
            {
                prevSlide = GetPrevSlide(tempSlide);
                tempSlide.Delete();
            }

            if (prevSlide.Shapes.Count == 0)
                return;

            PowerPoint.Slide addedSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(prevSlide.SlideIndex + 1, prevSlide.CustomLayout);
            //PowerPoint.Slide addedSlide = prevSlide.Duplicate()[1];
            addedSlide.Name = "PPTLabsZoomOut" + GetTimestamp(DateTime.Now);

            Globals.ThisAddIn.Application.ActiveWindow.Selection.Unselect();
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(currentSlide.SlideIndex);
            String tempFileName = "";
            foreach (PowerPoint.Shape sh in prevSlide.Shapes)
            {
                sh.Copy();
                sh.Name += GetTimestamp(DateTime.Now);
                tempFileName = Path.GetTempFileName();
                PowerPoint.Shape tempPicture1 = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                tempPicture1.Export(tempFileName, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                tempPicture1.LockAspectRatio = Office.MsoTriState.msoFalse;
                tempPicture1.Left = sh.Left;
                tempPicture1.Top = sh.Top;
                tempPicture1.Width = sh.Width;
                tempPicture1.Height = sh.Height;
                tempPicture1.Name = sh.Name;
                //tempPicture1.Delete();

                PowerPoint.Shape tempPicture2 = currentSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, sh.Left, sh.Top, sh.Width, sh.Height);
                tempPicture2.Name = sh.Name;
                tempPicture2.Select(Office.MsoTriState.msoFalse);
                sh.Name = sh.Name.Substring(0, sh.Name.Length - 18);
            }
            PowerPoint.Selection sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            PowerPoint.Shape shapeGroup = null;
            if (sel.ShapeRange.Count > 1)
            {
                shapeGroup = sel.ShapeRange.Group();
            }
            else
            {
                shapeGroup = sel.ShapeRange[1];
            }

            shapeGroup.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapeGroup.Left = shape.Left;
            shapeGroup.Top = shape.Top;
            shapeGroup.Width = shape.Width;
            shapeGroup.Height = shape.Height;
            //shapeGroup.Name = "PPTZoomInShape" + GetTimestamp(DateTime.Now);
            shape.Visible = Office.MsoTriState.msoFalse;

            //PowerPoint.ShapeRange range = null;
            List<PowerPoint.Shape> shapesToZoom = new List<PowerPoint.Shape>();
            if (shapeGroup.Type == Office.MsoShapeType.msoGroup)
            {
                //range = shapeGroup.Ungroup();
                foreach (PowerPoint.Shape tmp in shapeGroup.GroupItems)
                {
                    shapesToZoom.Add(tmp);
                }
            }
            else
            {
                shapesToZoom.Add(shapeGroup);
            }

            PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
            PowerPoint.Effect effectDisappear = null;
            PowerPoint.Effect effectMotion = null;
            PowerPoint.Effect effectResize = null;

            shapeGroup.Copy();
            PowerPoint.Shape shapePicture = currentSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            shapePicture.LockAspectRatio = Office.MsoTriState.msoFalse;
            shapePicture.Left = shape.Left;
            shapePicture.Top = shape.Top;
            shapePicture.Width = shape.Width;
            shapePicture.Height = shape.Height;

            tempFileName = Path.GetTempFileName();
            Properties.Resources.Indicator.Save(tempFileName);
            PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
            indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
            indicatorShape.Top = 0;
            indicatorShape.Width = 120;
            indicatorShape.Height = 84;
            indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
            effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;

            foreach (PowerPoint.Shape sh in addedSlide.Shapes)
            {
                PowerPoint.Shape matchingShape = null;
                foreach (PowerPoint.Shape tmpShape in shapesToZoom)
                {
                    if (sh.Name.Equals(tmpShape.Name))
                    {
                        matchingShape = tmpShape;
                        break;
                    }
                }
                if (matchingShape == null)
                {
                    if (!sh.Equals(indicatorShape))
                        sh.Delete();
                    continue;
                }
                    
                float initialX = (sh.Left + (sh.Width) / 2);
                float finalX = (matchingShape.Left + (matchingShape.Width) / 2);
                float initialY = (sh.Top + (sh.Height) / 2);
                float finalY = (matchingShape.Top + (matchingShape.Height) / 2);

                float initialWidth = sh.Width;
                float finalWidth = matchingShape.Width;
                float initialHeight = sh.Height;
                float finalHeight = matchingShape.Height;
                //float initialRotation = groupShape.Rotation;
                //float finalRotation = 0;

                effectMotion = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                PowerPoint.AnimationBehavior motion = effectMotion.Behaviors[1];
                effectMotion.Timing.Duration = defaultDuration;
                float point1X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                float point1Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                float point2X = ((finalX - initialX) / 2f) / presentation.PageSetup.SlideWidth;
                float point2Y = ((finalY - initialY) / 2f) / presentation.PageSetup.SlideHeight;
                float point3X = (finalX - initialX) / presentation.PageSetup.SlideWidth;
                float point3Y = (finalY - initialY) / presentation.PageSetup.SlideHeight;

                motion.MotionEffect.Path = "M 0 0 C " + point1X + " " + point1Y + " " + point2X + " " + point2Y + " " + point3X + " " + point3Y + " E";
                effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;

                effectResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];
                effectResize.Timing.Duration = defaultDuration;
                resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;
            }

            shapeGroup.Delete();
            indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
            addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
            addedSlide.SlideShowTransition.AdvanceTime = 0;
            addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;

            if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
            {
                foreach (PowerPoint.Shape sh in addedSlide.NotesPage.Shapes)
                {
                    if (sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        sh.TextEffect.Text = "";
                }
            }

            foreach (PowerPoint.Shape sh in addedSlide.Shapes)
            {
                if (sh.Type == Office.MsoShapeType.msoMedia)
                    sh.Delete();
            }
            //if (currentSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFade && currentSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFadeSmoothly)
            //    currentSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
            currentSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
            currentSlide.SlideShowTransition.Duration = 0.25f;
            Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
            AddAckSlide();
        }
        public void AddZoomOutButtonClick(Office.IRibbonControl control)
        {
            //System.Windows.Forms.MessageBox.Show("This feature is coming soon!                ", "Coming Soon");

            //Get References of current and next slides
            PowerPoint.Slide currentSlide = GetCurrentSlide();
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange.Count == 0)
                return;

            PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Shape shape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
            DeleteShapeAnnimations(currentSlide, shape);
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
            {
                if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Text = "";
                }
            }
            shape.Rotation = 0;

            if (currentSlide != null && currentSlide.SlideIndex != 1)
            {
                if (!backgroundZoomChecked)
                {
                    ZoomOutOmitBackground(currentSlide, shape);
                }
                else
                {
                    PowerPoint.Slide prevSlide = GetPrevSlide(currentSlide);
                    PowerPoint.Slide tempSlide = prevSlide;
                    while (prevSlide.Name.Contains("PPTLabsZoomOut") && prevSlide.SlideIndex > 1)
                    {
                        prevSlide = GetPrevSlide(tempSlide);
                        tempSlide.Delete();
                    }


                    String tempFileName = Path.GetTempFileName();
                    prevSlide.Export(tempFileName, "PNG");
                    shape.Fill.UserPicture(tempFileName);
                    shape.Line.Visible = Office.MsoTriState.msoFalse;
                    shape.Name = "PPTZoomOutShape" + GetTimestamp(DateTime.Now);

                    PowerPoint.Slide addedSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides.AddSlide(currentSlide.SlideIndex + 1, currentSlide.CustomLayout);
                    addedSlide.Name = "PPTLabsZoomOut" + GetTimestamp(DateTime.Now);
                    addedSlide.MoveTo(currentSlide.SlideIndex);

                    currentSlide.Copy();
                    PowerPoint.Shape zoomShape = addedSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                    zoomShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                    zoomShape.Left = 0;
                    zoomShape.Top = 0;
                    zoomShape.Width = presentation.PageSetup.SlideWidth;
                    zoomShape.Height = presentation.PageSetup.SlideHeight;
                    shape.Copy();
                    PowerPoint.Shape zoomCopyShape = addedSlide.Shapes.Paste()[1];

                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
                    zoomCopyShape.Select();
                    zoomShape.Select(Office.MsoTriState.msoFalse);
                    PowerPoint.ShapeRange selection = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                    PowerPoint.Shape groupShape = selection.Group();

                    //PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
                    //PowerPoint.Effect effectAppear = sequence.AddEffect(shape, PowerPoint.MsoAnimEffect.msoAnimEffectZoom, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                    //effectAppear.Timing.Duration = 0.25f;

                    PowerPoint.Sequence sequence = addedSlide.TimeLine.MainSequence;
                    groupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                    groupShape.Left = (presentation.PageSetup.SlideWidth / 2) - ((shape.Left + (shape.Width) / 2) * presentation.PageSetup.SlideWidth / shape.Width);
                    groupShape.Top = (presentation.PageSetup.SlideHeight / 2) - ((shape.Top + (shape.Height) / 2) * presentation.PageSetup.SlideHeight / shape.Height);
                    groupShape.Width = presentation.PageSetup.SlideWidth * presentation.PageSetup.SlideWidth / shape.Width;
                    groupShape.Height = presentation.PageSetup.SlideHeight * presentation.PageSetup.SlideHeight / shape.Height;
                    //groupShape.Rotation = -1 * shape.Rotation;

                    groupShape.Left += (0 - zoomCopyShape.Left);
                    groupShape.Top += (0 - zoomCopyShape.Top);

                    groupShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    //groupShape.Ungroup();
                    //zoomCopyShape.Delete();
                    PowerPoint.Effect effectDisappear = null;

                    tempFileName = Path.GetTempFileName();
                    Properties.Resources.Indicator.Save(tempFileName);
                    PowerPoint.Shape indicatorShape = addedSlide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, presentation.PageSetup.SlideWidth - 120, 0, 120, 84);
                    indicatorShape.Left = presentation.PageSetup.SlideWidth - 120;
                    indicatorShape.Top = 0;
                    indicatorShape.Width = 120;
                    indicatorShape.Height = 84;
                    indicatorShape.Name = "PPIndicator" + GetTimestamp(DateTime.Now);
                    effectDisappear = sequence.AddEffect(indicatorShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                    effectDisappear.Exit = Office.MsoTriState.msoTrue;
                    effectDisappear.Timing.Duration = 0;

                    float initialX = (groupShape.Left + (groupShape.Width) / 2);
                    float finalX = presentation.PageSetup.SlideWidth / 2;
                    float initialY = (groupShape.Top + (groupShape.Height) / 2);
                    float finalY = presentation.PageSetup.SlideHeight / 2;

                    float initialWidth = groupShape.Width;
                    float finalWidth = presentation.PageSetup.SlideWidth;
                    float initialHeight = groupShape.Height;
                    float finalHeight = presentation.PageSetup.SlideHeight;
                    float initialRotation = groupShape.Rotation;
                    //float finalRotation = 0;

                    int numFrames = 10;

                    float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
                    float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
                    //float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
                    float incrementLeft = (finalX - initialX) / numFrames;
                    float incrementTop = (finalY - initialY) / numFrames;

                    PowerPoint.Shape lastShape = groupShape;
                    for (int i = 1; i <= numFrames; i++)
                    {
                        PowerPoint.Shape dupShape = groupShape.Duplicate()[1];
                        if (i != 1)
                            sequence[sequence.Count].Delete();

                        dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                        dupShape.Left = groupShape.Left;
                        dupShape.Top = groupShape.Top;
                        //dupShape.Rotation = groupShape.Rotation;

                        if (incrementWidth != 0.0f)
                        {
                            dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                        }

                        if (incrementHeight != 0.0f)
                        {
                            dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                        }

                        //if (incrementRotation != 0.0f)
                        //{
                        //    dupShape.Rotation += (incrementRotation * i);
                        //}

                        if (incrementLeft != 0.0f)
                        {
                            dupShape.Left += (incrementLeft * i);
                        }

                        if (incrementTop != 0.0f)
                        {
                            dupShape.Top += (incrementTop * i);
                        }

                        PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        //appear.Timing.Duration = 0.005f;
                        appear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                        PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                        disappear.Exit = Office.MsoTriState.msoTrue;
                        //disappear.Timing.Duration = 0.005f;
                        disappear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                        lastShape = dupShape;
                    }

                    indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                    addedSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
                    addedSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
                    addedSlide.SlideShowTransition.AdvanceTime = 0;
                    addedSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;

                    if (addedSlide.HasNotesPage == Office.MsoTriState.msoTrue)
                    {
                        foreach (PowerPoint.Shape sh in addedSlide.NotesPage.Shapes)
                        {
                            if (sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                sh.TextEffect.Text = "";
                        }
                    }

                    foreach (PowerPoint.Shape sh in addedSlide.Shapes)
                    {
                        if (sh.Type == Office.MsoShapeType.msoMedia)
                            sh.Delete();
                    }
                    //if (currentSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFade && currentSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFadeSmoothly)
                    //    currentSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
                    currentSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
                    currentSlide.SlideShowTransition.Duration = 0.25f;
                    Globals.ThisAddIn.Application.ActiveWindow.View.GotoSlide(addedSlide.SlideIndex);
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                    AddAckSlide();
                }
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please select the correct slide", "Unable to Add Animations");
            }
        }

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            var selection = Globals.ThisAddIn.Application.
                ActiveWindow.Selection;
            CropShape(ref selection);
        }

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
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAnimation);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadAnimation);
            }
            catch (Exception e)
            {
                LogException(e, "GetReloadAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Spotlight);
            }
            catch (Exception e)
            {
                LogException(e, "GetSpotlightImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadSpotlight);
            }
            catch (Exception e)
            {
                LogException(e, "GetReloadSpotlightImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHighlightBulletsImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Bullets);
            }
            catch (Exception e)
            {
                LogException(e, "GetHighlightBulletsImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomInImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomIn);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomInImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomOutImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomOut);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomOutImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetMagnifyImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Magnify);
            }
            catch (Exception e)
            {
                LogException(e, "GetMagnifyImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetMagnifyContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.MagnifyContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetMagnifyContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetCropShapeImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.CutOutShape);
            }
            catch (Exception e)
            {
                LogException(e, "GetCropShapeImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAboutImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.About);
            }
            catch (Exception e)
            {
                LogException(e, "GetAboutImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetHelpImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Help);
            }
            catch (Exception e)
            {
                LogException(e, "GetHelpImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetFeedbackImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Feedback);
            }
            catch (Exception e)
            {
                LogException(e, "GetFeedbackImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAudio);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAudioImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetRemoveAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RemoveAudio);
            }
            catch (Exception e)
            {
                LogException(e, "GetRemoveAudioImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddCaption);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddCaptionsImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetRemoveCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RemoveCaption);
            }
            catch (Exception e)
            {
                LogException(e, "GetRemoveCaptionsImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetAddAudioContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddNarrationContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAudioContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetPreviewNarrationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.SpeakTextContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetPreviewNarrationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetInSlideAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.InSlideAnimation);
            }
            catch (Exception e)
            {
                LogException(e, "GetInSlideAnimationImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddAnimationContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetReloadAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ReloadAnimationContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetReloadAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetAddSpotlightContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.AddSpotlightContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetAddSpotlightContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetEditNameContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.EditNameContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetEditNameContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetInSlideAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.InSlideContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetInSlideAnimationContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomInContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomInContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomInContextImage");
                throw;
            }
        }
        public System.Drawing.Bitmap GetZoomOutContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ZoomOutContext);
            }
            catch (Exception e)
            {
                LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        //Duration Callbacks
        public void OnChangeDuration(Office.IRibbonControl control, String text)
        {
            try
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
            catch (Exception e)
            {
                LogException(e, "OnChangeDuration");
                throw;
            }
        }
        public String OnGetDurationText(Office.IRibbonControl control)
        {
            try
            {
                return defaultDuration.ToString();
            }
            catch (Exception e)
            {
                LogException(e, "OnGetDurationText");
                throw;
            }
        }

        //Checkbox Callbacks
        public void AnimationStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                if (pressed)
                {
                    frameAnimationChecked = true;
                }
                else
                {
                    frameAnimationChecked = false;
                }
            }
            catch (Exception e)
            {
                LogException(e, "AnimationStyleChanged");
                throw;
            }
        }
        public bool AnimationStyleGetPressed(Office.IRibbonControl control)
        {
            try
            {
                return frameAnimationChecked;
            }
            catch (Exception e)
            {
                LogException(e, "AnimationStyleGetPressed");
                throw;
            }
        }
        public void ZoomStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                if (pressed)
                {
                    backgroundZoomChecked = true;
                }
                else
                {
                    backgroundZoomChecked = false;
                }
            }
            catch (Exception e)
            {
                LogException(e, "ZoomStyleChanged");
                throw;
            }
        }
        public bool ZoomStyleGetPressed(Office.IRibbonControl control)
        {
            try
            {
                return backgroundZoomChecked;
            }
            catch (Exception e)
            {
                LogException(e, "ZoomStyleGetPressed");
                throw;
            }
        }
        //public void SpotlightDeleteOptionChanged(Office.IRibbonControl control, bool pressed)
        //{
        //    try
        //    {
        //        if (pressed)
        //        {
        //            spotlightDelete = true;
        //        }
        //        else
        //        {
        //            spotlightDelete = false;
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        LogException(e, "SpotlightDeleteOptionChanged");
        //        throw;
        //    }
        //}
        //public bool SpotlightDeleteOptionGetPressed(Office.IRibbonControl control)
        //{
        //    try
        //    {
        //        return spotlightDelete;
        //    }
        //    catch (Exception e)
        //    {
        //        LogException(e, "SpotlightDeleteOptionGetPressed");
        //        throw;
        //    }
        //}


        //Transparency Callbacks
        public void OnChangeTransparency(Office.IRibbonControl control, String text)
        {
            try
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
            catch (Exception e)
            {
                LogException(e, "OnChangeTransparency");
                throw;
            }
        }
        public String OnGetTransparency(Office.IRibbonControl control)
        {
            try
            {
                return (defaultTransparency * 100).ToString() + "%";
            }
            catch (Exception e)
            {
                LogException(e, "OnGetTransparency");
                throw;
            }
        }

        //Spotlight Edges Callbacks
        public int OnGetItemCountSpotlight(Office.IRibbonControl control)
        {
            try
            {
                return softEdgesMapping.Count;
            }
            catch (Exception e)
            {
                LogException(e, "OnGetItemCountSpotlight");
                throw;
            }
        }
        public String OnGetItemLabelSpotlight(Office.IRibbonControl control, int index)
        {
            try
            {
                String[] keys = softEdgesMapping.Keys.ToArray();
                return keys[index];
            }
            catch (Exception e)
            {
                LogException(e, "OnGetItemLabelSpotlight");
                throw;
            }
        }
        public void OnSelectItemSpotlight(Office.IRibbonControl control, String selectedId, int selectedIndex)
        {
            try
            {
                String[] keys = softEdgesMapping.Keys.ToArray();
                defaultSoftEdges = softEdgesMapping[keys[selectedIndex]];
            }
            catch (Exception e)
            {
                LogException(e, "OnSelectItemSpotlight");
                throw;
            }
        }
        public int OnGetSelectedItemIndexSpotlight(Office.IRibbonControl control)
        {
            try
            {
                float[] values = softEdgesMapping.Values.ToArray();
                return Array.IndexOf(values, defaultSoftEdges);
            }
            catch (Exception e)
            {
                LogException(e, "OnGetSelectedItemIndexSpotlight");
                throw;
            }
        }

        //Control Enabled Callbacks
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
        public bool OnGetEnabledAddInSlide(Office.IRibbonControl control)
        {
            return inSlideEnabled;
        }
        public bool OnGetEnabledZoomButton(Office.IRibbonControl control)
        {
            return zoomButtonEnabled;
        }

        //Edit Name Callbacks
        public void NameEditBtnClick(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                Form1 editForm = new Form1(this, selectedShape.Name);
                editForm.Show();
            }
            catch (Exception e)
            {
                LogException(e, "NameEditBtnClick");
                throw;
            }
        }
        public void ShapeNameEdited(String newName)
        {
            try
            {
                PowerPoint.Shape selectedShape = (PowerPoint.Shape)Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = newName;
            }
            catch (Exception e)
            {
                LogException(e, "ShapeNameEdited");
                throw;
            }
        }

        #region NotesToAudio Button Callbacks
        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }

        public void RemoveAudioClick(Office.IRibbonControl control)
        {
            if (_allSlides)
            {
                NotesToAudio.RemoveAudioFromAllSlides();
            }
            else
            {
                NotesToAudio.RemoveAudioFromCurrentSlide();
            }
        }

        public void AddAudioClick(Office.IRibbonControl control)
        {
            if (_allSlides)
            {
                NotesToAudio.EmbedAllSlideNotes();
            }
            else
            {
                NotesToAudio.EmbedCurrentSlideNotes();
            }

            PreviewAnimationsIfChecked();
        }

        public void ContextAddAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.EmbedCurrentSlideNotes();
            PreviewAnimationsIfChecked();
        }
        #endregion

        #region NotesToCaptions Button Callbacks

        public void AddCaptionClick(Office.IRibbonControl control)
        {
            if (_captionsAllSlides)
            {
                NotesToCaptions.EmbedCaptionsOnAllSlides();
            }
            else
            {
                NotesToCaptions.EmbedCaptionsOnCurrentSlide();
            }
        }

        public void RemoveCaptionClick(Office.IRibbonControl control)
        {
            if (_captionsAllSlides)
            {
                NotesToCaptions.RemoveCaptionsFromAllSlides();
            }
            else
            {
                NotesToCaptions.RemoveCaptionsFromCurrentSlide();
            }
        }

        public void ContextReplaceAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.ReplaceSelectedAudio();
        }

        #endregion

        #region NotesToAudio/Caption Helpers

        private void PreviewAnimationsIfChecked()
        {
            if (_previewCurrentSlide)
            {
                NotesToAudio.PreviewAnimations();
            }
        }

        public void AllSlidesChecked(Office.IRibbonControl control, bool pressed)
        {
            _allSlides = pressed;
        }

        public void PreviewCurrentSlideChecked(Office.IRibbonControl control, bool pressed)
        {
            _previewCurrentSlide = pressed;
        }

        public void AllSlidesCaptionsChecked(Office.IRibbonControl control, bool pressed)
        {
            _captionsAllSlides = pressed;
        }

        #region Dropdown Index/Label Handlers

        public int DefaultVoiceSelectedIndex(Office.IRibbonControl control)
        {
            return _voiceSelected;
        }

        public int DefaultVoicePickerCount(Office.IRibbonControl control)
        {
            return _voiceNames.Count();
        }

        public string DefaultVoicePickerLabel(Office.IRibbonControl control, int index)
        {
            return _voiceNames.ToArray()[index];
        }
        #endregion

        #region Dropdown Selection Handlers
        public void DefaultVoiceSelectionChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            _voiceSelected = selectedIndex;
            SetCoreVoicesToSelections();
        }

        #endregion

        private void SetCoreVoicesToSelections()
        {
            string defaultVoice = GetSelectedVoiceOrNull();
            NotesToAudio.SetDefaultVoice(defaultVoice);
        }

        private string GetSelectedVoiceOrNull()
        {
            string selectedVoice = null;
            try
            {
                selectedVoice = _voiceNames.ToArray()[_voiceSelected];
            }
            catch (IndexOutOfRangeException)
            {
                // No voices are installed.
                // (It should be impossible for the index to be out of range otherwise.)
            }
            return selectedVoice;
        }

        #endregion

        #endregion

        #region Fit To Slide | Fit To Width | Fit To Height

        private const int SelectedShapeIndex = 1;
        private const int TopMost = 0;
        private const int LeftMost = 0;

        public void FitToWidthClick(Office.IRibbonControl control)
        {
            DoFitToWidth();
        }

        public void FitToHeightClick(Office.IRibbonControl control)
        {
            DoFitToHeight();
        }

        private void DoFitToHeight()
        {
            var pageSetup = GetPageSetup();
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape.Height, selectedShape.Width);
            //fit to height
            selectedShape.Height = pageSetup.SlideHeight;
            selectedShape.Width = selectedShape.Height / shapeSizeRatio;
            //move to centre
            selectedShape.Left = (pageSetup.SlideWidth - selectedShape.Width) / 2;
            selectedShape.Top = TopMost;
        }

        private void DoFitToWidth()
        {
            var pageSetup = GetPageSetup();
            var selectedShape = GetSelectedShape();
            float shapeSizeRatio = GetSizeRatio(selectedShape.Height, selectedShape.Width);
            //fit to width
            selectedShape.Width = pageSetup.SlideWidth;
            selectedShape.Height = selectedShape.Width * shapeSizeRatio;
            //move to middle
            selectedShape.Top = (pageSetup.SlideHeight - selectedShape.Height) / 2;
            selectedShape.Left = LeftMost;
        }

        private float GetSizeRatio(float height, float width)
        {
            return height / width;
        }

        private PowerPoint.PageSetup GetPageSetup()
        {
            return Globals.ThisAddIn.Application.ActivePresentation.PageSetup;
        }

        private PowerPoint.Shape GetSelectedShape()
        {
            return Globals.ThisAddIn.Application.
                   ActiveWindow.Selection.ShapeRange[SelectedShapeIndex];
        }

        public System.Drawing.Bitmap GetFitToWidthImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.FitToWidth);
            }
            catch (Exception e)
            {
                LogException(e, "GetFitToWidthImage");
                throw;
            }
        }

        public System.Drawing.Bitmap GetFitToHeightImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.FitToHeight);
            }
            catch (Exception e)
            {
                LogException(e, "GetFitToHeightImage");
                throw;
            }
        }

        #endregion

        #region feature: Crop to Shape

        private void CropShape(ref PowerPoint.Selection selection)
        {
            if (IsValidSelectionForCropToShape(ref selection))
            {
                var shape = FormShapeForCropToShape(ref selection);
                string path = GetPathToStore();
                TakeScreenshotWithoutShape(ref shape, path);
                FillInShapeWithScreenshot(ref shape, path);
                RemoveOverlapItems(shape);
                ConvertToPicture(ref shape);
            }
            else
            {
                MessageBox.Show("Please select a shape", "Unable to crop");
            }
        }

        private string GetPathToStore()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments);
            return path + "\\currentSlide.png";
        }

        private void ConvertToPicture(ref PowerPoint.Shape shape)
        {
            shape.Copy();
            float x = shape.Left;
            float y = shape.Top;
            shape.Delete();
            var pic = GetCurrentSlide().Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            pic.Left = x;
            pic.Top = y;
        }

        private bool IsFirstOneOverlapWithSecond(PowerPoint.Shape first, PowerPoint.Shape second)
        {
            if (first == second)
                return false;
            if (!(first.Left > second.Left + second.Width ||
                  first.Left + first.Width < second.Left ||
                  first.Top > second.Top + second.Height ||
                  first.Top + first.Height < second.Top))
            {
                return true;
            }
            return false;
        }

        private void RemoveOverlapItems(PowerPoint.Shape shape)
        {
            var shapeRange = GetCurrentSlide().Shapes;
            for (int i = 1; i < shapeRange.Count; )
            {
                if (IsFirstOneOverlapWithSecond(shapeRange[i], shape))
                    shapeRange[i].Delete();
                else
                    i++;
            }
        }

        private void TakeScreenshotWithoutShape(ref PowerPoint.Shape shape, string fileToStore)
        {
            shape.Cut();
            TakeScreenshot(fileToStore);
            shape = GetCurrentSlide().Shapes.Paste()[1];
        }

        private void TakeScreenshot(string fileToStore)
        {
            GetCurrentSlide().Export(fileToStore, "PNG");
        }

        private void FillInShapeWithScreenshot(ref PowerPoint.Shape shape, string fileToStore)
        {
            var fillEffect = shape.Fill;
            fillEffect.UserTextured(fileToStore);
            fillEffect.TextureOffsetX = -shape.Left;
            fillEffect.TextureOffsetY = -shape.Top;
            AdjustFillEffect(ref shape);
            shape.Line.Visible = Office.MsoTriState.msoFalse;
        }

        private PowerPoint.Shape FormShapeForCropToShape(ref PowerPoint.Selection selection)
        {
            var range = UngroupAll(selection.ShapeRange);
            PowerPoint.Shape shape;
            if (range.Count > 1)
            {
                shape = range.Group();
            }
            else
            {
                shape = range[1];
            }
            return shape;
        }

        private bool IsValidSelectionForCropToShape(ref PowerPoint.Selection selection)
        {
            return selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes
                        && selection.ShapeRange.Count >= 1
                        && IsShape(selection);
        }

        private PowerPoint.ShapeRange UngroupAll(PowerPoint.ShapeRange range)
        {
            range.Cut();
            range = GetCurrentSlide().Shapes.Paste();
            List<string> ungroupedShapes = new List<string>();
            Queue<PowerPoint.Shape> selectedShapes = new Queue<PowerPoint.Shape>();

            foreach (var item in range)
            {
                selectedShapes.Enqueue(item as PowerPoint.Shape);
            }
            while (selectedShapes.Count != 0)
            {
                var shape = selectedShapes.Dequeue();
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subRange = shape.Ungroup();
                    foreach (var item in subRange)
                    {
                        selectedShapes.Enqueue(item as PowerPoint.Shape);
                    }
                }
                else
                {
                    ungroupedShapes.Add(shape.Name);
                }
            }
            return GetCurrentSlide().Shapes.Range(ungroupedShapes.ToArray());
        }

        private void AdjustFillEffect(ref PowerPoint.Shape shape)
        {
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                var range = shape.Ungroup();
                foreach (var o in range)
                {
                    var sh = o as PowerPoint.Shape;
                    var tmpFillEffect = sh.Fill;
                    tmpFillEffect.TextureOffsetX = -sh.Left;
                    tmpFillEffect.TextureOffsetY = -sh.Top;
                }
                shape = range.Group();
            }
        }

        private bool IsShape(PowerPoint.Selection sel)
        {
            var shapeRange = sel.ShapeRange;
            foreach (var o in shapeRange)
            {
                var shape = o as PowerPoint.Shape;
                if (shape.Type != Office.MsoShapeType.msoAutoShape
                    && shape.Type != Office.MsoShapeType.msoFreeform
                    && shape.Type != Office.MsoShapeType.msoGroup)
                {
                    return false;
                }
            }
            return true;
        }

        public System.Drawing.Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.CutOutShapeMenu);
            }
            catch (Exception e)
            {
                LogException(e, "GetCutOutShapeMenuImage");
                throw;
            }
        }

        #endregion

        #region Helpers

        //Spotlight Helpers
        private void AddSpotlightEffect(PowerPoint.Slide addedSlide, List<PowerPoint.Shape> spotlightShapes)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                PowerPoint.Shape rectangleShape = addedSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, (-1 * defaultSoftEdges), (-1 * defaultSoftEdges), (presentation.PageSetup.SlideWidth + (2.0f * defaultSoftEdges)), (presentation.PageSetup.SlideHeight + (2.0f * defaultSoftEdges)));
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
                pictureShape.Left = -1 * defaultSoftEdges;
                pictureShape.Top = -1 * defaultSoftEdges;
                pictureShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                float incrementWidth = (2.0f * defaultSoftEdges) / pictureShape.Width;
                float incrementHeight = (2.0f * defaultSoftEdges) / pictureShape.Height;

                pictureShape.SoftEdge.Radius = defaultSoftEdges;
                //pictureShape.SoftEdge.Type = Office.MsoSoftEdgeType.msoSoftEdgeType4;
                //pictureShape.ScaleWidth((1.0f + incrementWidth), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                //pictureShape.ScaleHeight((1.0f + incrementHeight), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                //pictureShape.Width = presentation.PageSetup.SlideWidth + (2.0f * defaultSoftEdges);
                //pictureShape.Height = presentation.PageSetup.SlideHeight + (2.0f * defaultSoftEdges);
                //pictureShape.Left = (-1 * defaultSoftEdges);
                //pictureShape.Top = (-1 * defaultSoftEdges);
                pictureShape.Name = "SpotlightShape1";
            }
            catch (Exception e)
            {
                LogException(e, "AddSpotlightEffect");
                throw;
            }
        }

        //AutoAnimate Helpers
        private void AddCompleteAutoMotion(PowerPoint.Slide currentSlide, PowerPoint.Slide nextSlide)
        {
            try
            {
                //Create containers to store information on matching shapes
                PowerPoint.Shape[] shapes1;
                PowerPoint.Shape[] shapes2;
                int[] shapeIDs;

                if (GetMatchingShapeDetails(currentSlide, nextSlide, out shapes1, out shapes2, out shapeIDs))
                {
                    //If an identical object exists
                    AboutForm progressForm = new AboutForm();
                    progressForm.Visible = true;
                    PowerPoint.Slide newSlide = PrepareAnimatedSlide(currentSlide, shapeIDs);
                    AddAnimationsToShapes(newSlide, shapes1, shapes2, shapeIDs);
                    //this.ribbon.ActivateTabMso("TabAnimations");
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
                    progressForm.Visible = false;
                }
                else
                {
                    //Display error message
                    System.Windows.Forms.MessageBox.Show("No matching Shapes were found on the next slide", "Animation Not Added");
                }
            }
            catch (Exception e)
            {
                LogException(e, "AddCompleteAutoMotion");
                throw;
            }
        }
        private void AddAnimationsToShapes(PowerPoint.Slide newSlide, PowerPoint.Shape[] shapes1, PowerPoint.Shape[] shapes2, int[] shapeIDs)
        {
            try
            {
                int count = 0;
                bool fadeFlag = false;
                PowerPoint.Sequence sequence = newSlide.TimeLine.MainSequence;
                PowerPoint.Effect effectMotion = null;
                PowerPoint.Effect effectResize = null;
                PowerPoint.Effect effectRotate = null;
                //PowerPoint.Effect effectFontResize = null;
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
                        if (DeleteShapeAnnimations(newSlide, sh))
                        {
                            effectFade = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectFade, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectFade.Exit = Office.MsoTriState.msoTrue;
                            effectFade.Timing.Duration = defaultDuration;
                            fadeFlag = true;
                        }
                        else
                        {
                            PowerPoint.Effect effectDisappear2 = null;
                            effectDisappear2 = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            effectDisappear2.Exit = Office.MsoTriState.msoTrue;
                            effectDisappear2.Timing.Duration = 0;
                        }
                    }
                }

                //Add animation effects to the duplicated objects
                foreach (PowerPoint.Shape sh in newSlide.Shapes)
                {
                    if (shapeIDs.Contains(sh.Id))
                    {
                        count = Array.IndexOf(shapeIDs, sh.Id);
                        if (count < shapeIDs.Count() && sh.Id == shapeIDs[count])
                        {
                            DeleteShapeAnnimations(newSlide, sh);
                            trigger = (count == 0 && fadeFlag) ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious : PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                            float finalX = (shapes2[count].Left + (shapes2[count].Width) / 2);
                            float initialX = (sh.Left + (sh.Width) / 2);
                            float finalY = (shapes2[count].Top + (shapes2[count].Height) / 2);
                            float initialY = (sh.Top + (sh.Height) / 2);

                            float finalRotation = shapes2[count].Rotation;
                            float initialRotation = sh.Rotation;

                            float finalWidth = shapes2[count].Width;
                            float initialWidth = sh.Width;
                            float finalHeight = shapes2[count].Height;
                            float initialHeight = sh.Height;
                            float finalFont = 0.0f;
                            float initialFont = 0.0f;
                            int numFrames = (int)(defaultDuration / 0.04f);
                            numFrames = (numFrames > 30) ? 30 : numFrames;

                            if (sh.HasTextFrame == Office.MsoTriState.msoTrue && (sh.TextFrame.HasText == Office.MsoTriState.msoTriStateMixed || sh.TextFrame.HasText == Office.MsoTriState.msoTrue) && sh.TextFrame.TextRange.Font.Size != shapes2[count].TextFrame.TextRange.Font.Size)
                            {
                                finalFont = shapes2[count].TextFrame.TextRange.Font.Size;
                                initialFont = sh.TextFrame.TextRange.Font.Size;
                            }

                            if ((frameAnimationChecked && (finalHeight != initialHeight || finalWidth != initialWidth))
                                || ((initialRotation != finalRotation || initialRotation % 90 != 0) && (finalHeight != initialHeight || finalWidth != initialWidth))
                                || finalFont != initialFont)
                            {
                                float incrementWidth = ((finalWidth / initialWidth) - 1.0f) / numFrames;
                                float incrementHeight = ((finalHeight / initialHeight) - 1.0f) / numFrames;
                                float incrementRotation = GetMinimumRotation(initialRotation, finalRotation) / numFrames;
                                float incrementLeft = (finalX - initialX) / numFrames;
                                float incrementTop = (finalY - initialY) / numFrames;
                                float incrementFont = (finalFont - initialFont) / numFrames;

                                PowerPoint.Shape lastShape = sh;
                                for (int i = 1; i <= numFrames; i++)
                                {
                                    PowerPoint.Shape dupShape = sh.Duplicate()[1];
                                    if (i != 1)
                                        sequence[sequence.Count].Delete();

                                    dupShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                                    dupShape.Left = sh.Left;
                                    dupShape.Top = sh.Top;

                                    if (incrementWidth != 0.0f)
                                    {
                                        dupShape.ScaleWidth((1.0f + (incrementWidth * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                    }

                                    if (incrementHeight != 0.0f)
                                    {
                                        dupShape.ScaleHeight((1.0f + (incrementHeight * i)), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                    }

                                    if (incrementRotation != 0.0f)
                                    {
                                        dupShape.Rotation += (incrementRotation * i);
                                    }

                                    if (incrementLeft != 0.0f)
                                    {
                                        dupShape.Left += (incrementLeft * i);
                                    }

                                    if (incrementTop != 0.0f)
                                    {
                                        dupShape.Top += (incrementTop * i);
                                    }

                                    if (incrementFont != 0.0f)
                                    {
                                        dupShape.TextFrame.TextRange.Font.Size += (incrementFont * i);
                                    }

                                    PowerPoint.Effect appear = sequence.AddEffect(dupShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    //appear.Timing.Duration = 0.005f;
                                    appear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

                                    PowerPoint.Effect disappear = sequence.AddEffect(lastShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                                    disappear.Exit = Office.MsoTriState.msoTrue;
                                    //disappear.Timing.Duration = 0.005f;
                                    disappear.Timing.TriggerDelayTime = ((defaultDuration / numFrames) * i);

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
                                    effectMotion.Timing.SmoothStart = Office.MsoTriState.msoFalse;
                                    effectMotion.Timing.SmoothEnd = Office.MsoTriState.msoFalse;
                                }

                                //Resize Effect
                                //if (sh.Type != Office.MsoShapeType.msoPlaceholder && sh.Type != Office.MsoShapeType.msoTextBox)
                                //{
                                if ((finalWidth != initialWidth) || (finalHeight != initialHeight))
                                {
                                    sh.LockAspectRatio = Office.MsoTriState.msoFalse;
                                    effectResize = sequence.AddEffect(sh, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, trigger);
                                    //PowerPoint.AnimationBehavior resize = effectResize.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeScale);
                                    PowerPoint.AnimationBehavior resize = effectResize.Behaviors[1];

                                    //float rotCos = (float)Math.Cos(degToRad(sh.Rotation));
                                    //float rotSin = (float)Math.Sin(degToRad(sh.Rotation));

                                    effectResize.Timing.Duration = defaultDuration;
                                    //sh.ScaleWidth((finalWidth / initialWidth), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                                    //sh.ScaleHeight((finalHeight / initialHeight), Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);

                                    resize.ScaleEffect.ByX = (finalWidth / initialWidth) * 100;
                                    resize.ScaleEffect.ByY = (finalHeight / initialHeight) * 100;

                                    //resize.ScaleEffect.ByX = (((finalWidth / initialWidth) * Math.Abs(rotCos)) + ((finalHeight / initialHeight) * Math.Abs(rotSin))) * 100;
                                    //resize.ScaleEffect.ByY = (((finalWidth / initialWidth) * Math.Abs(rotSin)) + ((finalHeight / initialHeight) * Math.Abs(rotCos))) * 100;
                                    trigger = PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                                }
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
                indicatorShape.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
            }
            catch (Exception e)
            {
                LogException(e, "AddAnimationsToShapes");
                throw;
            }
        }
        private PowerPoint.Slide PrepareAnimatedSlide(PowerPoint.Slide currentSlide, int[] shapeIDs)
        {
            try
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
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                //Delete non-identical shapes
                foreach (PowerPoint.Effect eff in newSlide.TimeLine.MainSequence)
                {
                    if ((eff.EffectType >= PowerPoint.MsoAnimEffect.msoAnimEffectPathCircle && eff.EffectType <= PowerPoint.MsoAnimEffect.msoAnimEffectPathRight) || eff.EffectType == PowerPoint.MsoAnimEffect.msoAnimEffectCustom)
                    {
                        //sh.Delete();
                        PowerPoint.AnimationBehavior motion = eff.Behaviors[1];
                        if (motion.Type == PowerPoint.MsoAnimType.msoAnimTypeMotion)
                        {
                            PowerPoint.Shape sh = eff.Shape;
                            string motionPath = motion.MotionEffect.Path.Trim();
                            if (motionPath.Last() < 'A' || motionPath.Last() > 'Z')
                                motionPath += " X";
                            string[] path = motionPath.Split(' ');
                            int count = path.Length;
                            float xVal = Convert.ToSingle(path[count - 3]);
                            float yVal = Convert.ToSingle(path[count - 2]);
                            sh.Left += (xVal * presentation.PageSetup.SlideWidth);
                            sh.Top += (yVal * presentation.PageSetup.SlideHeight);
                        }
                    }
                }

                //Manage Slide Transitions
                newSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
                //newSlide.SlideShowTransition.Duration = defaultDuration;
                newSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
                newSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoFalse;
                newSlide.SlideShowTransition.AdvanceTime = 0;
                if (newSlide.HasNotesPage == Office.MsoTriState.msoTrue)
                {
                    foreach (PowerPoint.Shape sh in newSlide.NotesPage.Shapes)
                    {
                        if (sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                            sh.TextEffect.Text = "";
                    }
                }

                foreach (PowerPoint.Shape sh in newSlide.Shapes)
                {
                    if (sh.Type == Office.MsoShapeType.msoMedia)
                        sh.Delete();
                }

                PowerPoint.Slide nextSlide = GetNextSlide(newSlide);
                if (nextSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFade && nextSlide.SlideShowTransition.EntryEffect != PowerPoint.PpEntryEffect.ppEffectFadeSmoothly)
                    nextSlide.SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectNone;
                if (nextSlide.Name.Contains("PPSlideStart") || nextSlide.Name.Contains("PPSlideMulti"))
                    nextSlide.Name = "PPSlideMulti" + GetTimestamp(DateTime.Now);
                else
                    nextSlide.Name = "PPSlideEnd" + GetTimestamp(DateTime.Now);

                AddAckSlide();

                return newSlide;
            }
            catch (Exception e)
            {
                LogException(e, "PrepareAnimatedSlide");
                throw;
            }
        }
        private bool GetMatchingShapeDetails(PowerPoint.Slide currentSlide, PowerPoint.Slide nextSlide, out PowerPoint.Shape[] shapes1, out PowerPoint.Shape[] shapes2, out int[] shapeIDs)
        {
            try
            {
                bool flag = false;
                int counter = 0;

                shapes1 = new PowerPoint.Shape[currentSlide.Shapes.Count];
                shapes2 = new PowerPoint.Shape[currentSlide.Shapes.Count];
                shapeIDs = new int[currentSlide.Shapes.Count];
                PowerPoint.Shape matchingShape;

                foreach (PowerPoint.Shape sh1 in currentSlide.Shapes)
                {
                    matchingShape = null;
                    foreach (PowerPoint.Shape sh2 in nextSlide.Shapes)
                    {
                        if (sh1.Id == sh2.Id && haveSameNames(sh1, sh2))
                        {
                            if (matchingShape == null)
                                matchingShape = sh2;
                            else
                            {
                                if (GetDistanceBetweenShapes(sh1, sh2) < GetDistanceBetweenShapes(sh1, matchingShape))
                                    matchingShape = sh2;
                            }
                            flag = true;
                        }
                    }
                    if (matchingShape == null)
                    {
                        foreach (PowerPoint.Shape sh2 in nextSlide.Shapes)
                        {
                            if (haveSameNames(sh1, sh2))
                            {
                                if (matchingShape == null)
                                    matchingShape = sh2;
                                else
                                {
                                    if (GetDistanceBetweenShapes(sh1, sh2) < GetDistanceBetweenShapes(sh1, matchingShape))
                                        matchingShape = sh2;
                                }
                                flag = true;
                            }
                        }
                    }
                    if (matchingShape != null)
                    {
                        shapes1[counter] = sh1;
                        shapes2[counter] = matchingShape;
                        shapeIDs[counter] = sh1.Id;
                        counter++;
                    }
                }
                return flag;
            }
            catch (Exception e)
            {
                LogException(e, "GetMatchingShapeDetails");
                throw;
            }
        }
        private float GetDistanceBetweenShapes(PowerPoint.Shape sh1, PowerPoint.Shape sh2)
        {
            float sh1CenterX = (sh1.Left + (sh1.Width / 2));
            float sh2CenterX = (sh2.Left + (sh2.Width / 2));
            float sh1CenterY = (sh1.Top + (sh1.Height / 2));
            float sh2CenterY = (sh2.Top + (sh2.Height / 2));
            float distSquared = (float)(Math.Pow((sh2CenterX - sh1CenterX), 2) +  Math.Pow((sh2CenterY - sh1CenterY), 2));
            return (float)(Math.Sqrt(distSquared));
        }
        private float GetMinimumRotation(float fromAngle, float toAngle)
        {
            try
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
            catch (Exception e)
            {
                LogException(e, "GetMinimumRotation");
                throw;
            }
        }
        private float Normalize(float i)
        {
            try
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
            catch (Exception e)
            {
                LogException(e, "Normalize");
                throw;
            }
        }
        private bool DeleteShapeAnnimations(PowerPoint.Slide slide, PowerPoint.Shape shape)
        {
            try
            {
                PowerPoint.Sequence sequence = slide.TimeLine.MainSequence;
                bool flag = true;
                for (int x = sequence.Count; x >= 1; x--)
                {
                    PowerPoint.Effect effect = sequence[x];
                    if (effect.Shape.Name == shape.Name && effect.Shape.Id == shape.Id)
                    {
                        if (effect.Exit == Office.MsoTriState.msoTrue)
                            flag = false;
                        effect.Delete();
                    }
                }
                return flag;
            }
            catch (Exception e)
            {
                LogException(e, "DeleteShapeAnimations");
                throw;
            }
        }
        //Other Helpers
        private void AddAckSlide()
        {
            try
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
            catch (Exception e)
            {
                LogException(e, "AddAckSlide");
                throw;
            }
        }
        private void Log(string logText, string type)
        {
            if (type.Equals("Info"))
                Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            else if (type.Equals("Error"))
                Trace.TraceError(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            else if (type.Equals("Warning"))
                Trace.TraceWarning(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
        }
        private void LogException(Exception e, string methodName)
        {
            Log(methodName + ": " + e.Message + ": " + e.StackTrace, "Error");
        }
        private PowerPoint.Slide GetCurrentSlide()
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                return Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            }
            catch (Exception e)
            {
                LogException(e, "GetCurrentSlide");
                throw;
            }
        }
        private PowerPoint.Slide GetNextSlide(PowerPoint.Slide currentSlide)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                int slideIndex = currentSlide.SlideIndex;
                return presentation.Slides[slideIndex + 1];
            }
            catch (Exception e)
            {
                LogException(e, "GetNextSlide");
                throw;
            }
        }
        private PowerPoint.Slide GetPrevSlide(PowerPoint.Slide currentSlide)
        {
            try
            {
                PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;
                int slideIndex = currentSlide.SlideIndex;
                return presentation.Slides[slideIndex - 1];
            }
            catch (Exception e)
            {
                LogException(e, "GetPrevSlide");
                throw;
            }
        }
        private bool haveSameNames(PowerPoint.Shape sh1, PowerPoint.Shape sh2)
        {
            try
            {
                String name1 = sh1.Name;
                String name2 = sh2.Name;

                if (name1.ToUpper().CompareTo(name2.ToUpper()) == 0)
                    return true;
                else
                    return false;
            }
            catch (Exception e)
            {
                LogException(e, "haveSameNames");
                throw;
            }
        }
        public String GetTimestamp(DateTime value)
        {
            try
            {
                return value.ToString("yyyyMMddHHmmssffff");
            }
            catch (Exception e)
            {
                LogException(e, "GetTimestamp");
                throw;
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
