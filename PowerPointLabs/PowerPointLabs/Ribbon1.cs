using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using PowerPointLabs.Models;
using PowerPointLabs.Views;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

// Follow these steps to enable the Ribbon (XML) item:

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
        private Office.IRibbonUI _ribbon;
        
        public bool FrameAnimationChecked = false;
        public bool BackgroundZoomChecked = true;
        public bool MultiSlideZoomChecked = false;
        public bool SpotlightDelete = true;
        public float DefaultDuration = 0.5f;
        
        public bool SpotlightEnabled = false;
        public bool InSlideEnabled = false;
        public bool ZoomButtonEnabled = false;
        public bool HighlightBulletsEnabled = true;
        public bool AddAutoMotionEnabled = true;
        public bool ReloadAutoMotionEnabled = true;
        public bool ReloadSpotlight = true;
        public bool RemoveCaptionsEnabled = true;
        public bool RemoveAudioEnabled = true;

        public bool EmbedAudioVisible = true;
        public bool RecorderPaneVisible = false;

        private bool _previewCurrentSlide;
        
        private List<string> _voiceNames;

        private int _voiceSelected;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("PowerPointLabs.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void RibbonLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;

            SetVoicesFromInstalledOptions();
            SetCoreVoicesToSelections();
        }

        public void RefreshRibbonControl(String controlId)
        {
            try
            {
                _ribbon.InvalidateControl(controlId);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "RefreshRibbonControl");
                throw;
            }
        }

        private void SetVoicesFromInstalledOptions()
        {
            var installedVoices = NotesToAudio.GetVoices().ToList();
            _voiceNames = installedVoices;
        }

        public void HighlightBulletsBackgroundButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kTextSelected;
                else
                    HighlightBulletsBackground.userSelection = HighlightBulletsBackground.HighlightBackgroundSelection.kNoneSelected;

                HighlightBulletsBackground.AddHighlightBulletsBackground();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsBackgroundButtonClick");
                throw;
            }
        }

        public void HighlightBulletsTextButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kTextSelected;
                else
                    HighlightBulletsText.userSelection = HighlightBulletsText.HighlightTextSelection.kNoneSelected;

                HighlightBulletsText.AddHighlightBulletsText();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsTextButtonClick");
                throw;
            }
        }

        public void AddInSlideAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AnimateInSlide.isHighlightBullets = false;
                AnimateInSlide.AddAnimationInSlide();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddInSlideAnimationButtonClick");
                throw;
            }
        }
        public void ReloadSpotlightButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Spotlight.ReloadSpotlightEffect();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ReloadSpotlightButtonClick");
                throw;
            }
        }
        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            try
            {
                Spotlight.AddSpotlightEffect();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        // Supertips Callbacks
        public string GetAddAnimationButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationButtonSupertip;
        }
        public string GetReloadButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ReloadButtonSupertip;
        }
        public string GetInSlideAnimateButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.InSlideAnimateButtonSupertip;
        }
        
        public string GetAddZoomInButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomInButtonSupertip;
        }
        public string GetAddZoomOutButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomOutButtonSupertip;
        }
        public string GetZoomToAreaButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ZoomToAreaButtonSupertip;
        }
        
        public string GetMoveCropShapeButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.MoveCropShapeButtonSupertip;
        }
        
        public string GetAddSpotlightButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddSpotlightButtonSupertip;
        }
        public string GetReloadSpotlightButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ReloadSpotlightButtonSupertip;
        }
        
        public string GetAddAudioButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddAudioButtonSupertip;
        }
        public string GetGenerateRecordButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.GenerateRecordButtonSupertip;
        }
        public string GetAddRecordButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddRecordButtonSupertip;
        }
        public string GetRemoveAudioButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAudioButtonSupertip;
        }
        
        public string GetAddCaptionsButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddCaptionsButtonSupertip;
        }
        public string GetRemoveCaptionsButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.RemoveCaptionsButtonSupertip;
        }
        
        public string GetHighlightBulletsTextButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsTextButtonSupertip;
        }
        public string GetHighlightBulletsBackgroundButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsBackgroundButtonSupertip;
        }
        
        public string GetColorPickerButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ColorPickerButtonSupertip;
        }
        
        public string GetCustomeShapeButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.CustomeShapeButtonSupertip;
        }
        
        public string GetHelpButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HelpButtonSupertip;
        }
        public string GetFeedbackButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.FeedbackButtonSupertip;
        }
        public string GetAboutButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AboutButtonSupertip;
        }


        //Button Click Callbacks        
        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimate.AddAutoAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddAnimationButtonClick");
                throw;
            }
        }
        public void ReloadButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoAnimate.ReloadAutoAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ReloadAnimationButtonClick");
                throw;
            }
        }
        public void ZoomBtnClick(Office.IRibbonControl control)
        {
            ZoomToArea.AddZoomToArea();
        }
        public void AboutButtonClick(Office.IRibbonControl control)
        {
            MessageBox.Show(TextCollection.AboutInfo, TextCollection.AboutInfoTitle);
        }
        public void HelpButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Process.Start(TextCollection.HelpDocumentUrl);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HelpButtonClick");
                throw;
            }
        }
        public void FeedbackButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Process.Start(TextCollection.FeedbackUrl);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "FeedbackButtonClick");
                throw;
            }
        }
        public void AddZoomInButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoZoom.AddDrillDownAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddZoomInButtonClick");
                throw;
            }
        }
        public void AddZoomOutButtonClick(Office.IRibbonControl control)
        {
            try
            {
                AutoZoom.AddStepBackAnimation();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AddZoomOutButtonClick");
                throw;
            }
        }

        public Bitmap GetAddAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAnimationImage");
                throw;
            }
        }
        public Bitmap GetReloadAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ReloadAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadAnimationImage");
                throw;
            }
        }
        public Bitmap GetSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Spotlight);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetSpotlightImage");
                throw;
            }
        }
        public Bitmap GetReloadSpotlightImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ReloadSpotlight);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadSpotlightImage");
                throw;
            }
        }
        public Bitmap GetHighlightBulletsTextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightText);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsTextImage");
                throw;
            }
        }
        public Bitmap GetHighlightBulletsBackgroundImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightBackground);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsBackgroundImage");
                throw;
            }
        }

        public Bitmap GetHighlightBulletsTextContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightTextContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsTextContextImage");
                throw;
            }
        }
        public Bitmap GetHighlightBulletsBackgroundContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightBackgroundContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHighlightBulletsBackgroundContextImage");
                throw;
            }
        }

        public Bitmap GetZoomInImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomIn);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomInImage");
                throw;
            }
        }

        public Bitmap GetZoomOutImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomOut);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomOutImage");
                throw;
            }
        }
        public Bitmap GetZoomToAreaImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomToArea);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomToAreaImage");
                throw;
            }
        }
        public Bitmap GetZoomToAreaContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomToAreaContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomToAreaContextImage");
                throw;
            }
        }
        public Bitmap GetCropShapeImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.CutOutShape);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetCropShapeImage");
                throw;
            }
        }
        public Bitmap GetAboutImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.About);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAboutImage");
                throw;
            }
        }
        public Bitmap GetHelpImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Help);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetHelpImage");
                throw;
            }
        }
        public Bitmap GetFeedbackImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.Feedback);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetFeedbackImage");
                throw;
            }
        }
        public Bitmap GetAddAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddAudio);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAudioImage");
                throw;
            }
        }
        public Bitmap GetRemoveAudioImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.RemoveAudio);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetRemoveAudioImage");
                throw;
            }
        }
        public Bitmap GetAddCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddCaption);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddCaptionsImage");
                throw;
            }
        }
        public Bitmap GetRemoveCaptionsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.RemoveCaption);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetRemoveCaptionsImage");
                throw;
            }
        }

        public Bitmap GetAddAudioContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddNarrationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAudioContextImage");
                throw;
            }
        }
        public Bitmap GetPreviewNarrationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.SpeakTextContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetPreviewNarrationContextImage");
                throw;
            }
        }
        public Bitmap GetInSlideAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.InSlideAnimation);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetInSlideAnimationImage");
                throw;
            }
        }
        public Bitmap GetAddAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddAnimationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddAnimationContextImage");
                throw;
            }
        }
        public Bitmap GetReloadAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ReloadAnimationContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetReloadAnimationContextImage");
                throw;
            }
        }
        public Bitmap GetAddSpotlightContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddSpotlightContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetAddSpotlightContextImage");
                throw;
            }
        }
        public Bitmap GetEditNameContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.EditNameContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetEditNameContextImage");
                throw;
            }
        }
        public Bitmap GetInSlideAnimationContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.InSlideContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetInSlideAnimationContextImage");
                throw;
            }
        }
        public Bitmap GetZoomInContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomInContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomInContextImage");
                throw;
            }
        }
        public Bitmap GetZoomOutContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.ZoomOutContext);
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        public void ZoomStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                BackgroundZoomChecked = pressed;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomStyleChanged");
                throw;
            }
        }
        public bool ZoomStyleGetPressed(Office.IRibbonControl control)
        {
            try
            {
                return BackgroundZoomChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomStyleGetPressed");
                throw;
            }
        }
        //Control Enabled Callbacks
        public bool OnGetEnabledSpotlight(Office.IRibbonControl control)
        {
            return SpotlightEnabled;
        }
        public bool OnGetEnabledReloadSpotlight(Office.IRibbonControl control)
        {
            return ReloadSpotlight;
        }
        public bool OnGetEnabledAddAutoMotion(Office.IRibbonControl control)
        {
            return AddAutoMotionEnabled;
        }
        public bool OnGetEnabledReloadAutoMotion(Office.IRibbonControl control)
        {
            return ReloadAutoMotionEnabled;
        }
        public bool OnGetEnabledAddInSlide(Office.IRibbonControl control)
        {
            return InSlideEnabled;
        }
        public bool OnGetEnabledZoomButton(Office.IRibbonControl control)
        {
            return ZoomButtonEnabled;
        }
        public bool OnGetEnabledHighlightBullets(Office.IRibbonControl control)
        {
            return HighlightBulletsEnabled;
        }
        public bool OnGetEnabledRemoveCaptions(Office.IRibbonControl control)
        {
            return RemoveCaptionsEnabled;
        }
        public bool OnGetEnabledRemoveAudio(Office.IRibbonControl control)
        {
            return RemoveAudioEnabled;
        }
        //Edit Name Callbacks
        public void NameEditBtnClick(Office.IRibbonControl control)
        {
            try
            {
                PowerPoint.Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                var editForm = new Form1(this, selectedShape.Name);
                editForm.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "NameEditBtnClick");
                throw;
            }
        }
        public void ShapeNameEdited(String newName)
        {
            try
            {
                PowerPoint.Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = newName;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ShapeNameEdited");
                throw;
            }
        }

        public void AutoAnimateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoAnimateDialogBox(DefaultDuration, FrameAnimationChecked);
                dialog.SettingsHandler += AnimationPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoAnimateDialogButtonPressed");
                throw;
            }
        }

        public void AnimationPropertiesEdited(float newDuration, bool newFrameChecked)
        {
            try
            {
                DefaultDuration = newDuration;
                FrameAnimationChecked = newFrameChecked;
                AnimateInSlide.defaultDuration = newDuration;
                AnimateInSlide.frameAnimationChecked = newFrameChecked;
                AutoAnimate.defaultDuration = newDuration;
                AutoAnimate.frameAnimationChecked = newFrameChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AnimationPropertiesEdited");
                throw;
            }
        }

        public void AutoZoomDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoZoomDialogBox(BackgroundZoomChecked, MultiSlideZoomChecked);
                dialog.SettingsHandler += ZoomPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoZoomDialogButtonPressed");
                throw;
            }
        }

        public void ZoomPropertiesEdited(bool backgroundChecked, bool multiSlideChecked)
        {
            try
            {
                BackgroundZoomChecked = backgroundChecked;
                MultiSlideZoomChecked = multiSlideChecked;
                AutoZoom.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.backgroundZoomChecked = backgroundChecked;
                ZoomToArea.multiSlideZoomChecked = multiSlideChecked;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "ZoomPropertiesEdited");
                throw;
            }
        }

        public void SpotlightDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new SpotlightDialogBox(Spotlight.defaultTransparency, Spotlight.defaultSoftEdges);
                dialog.SettingsHandler += SpotlightPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "SpotlightDialogButtonPressed");
                throw;
            }
        }

        public void SpotlightPropertiesEdited(float newTransparency, float newSoftEdge)
        {
            try
            {
                Spotlight.defaultTransparency = newTransparency;
                Spotlight.defaultSoftEdges = newSoftEdge;
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "SpotlightPropertiesEdited");
                throw;
            }
        }

        public void HighlightBulletsPropertiesEdited(Color newHighlightColor, Color newDefaultColor, Color newBackgroundColor)
        {
            try
            {
                HighlightBulletsText.highlightColor = newHighlightColor;
                HighlightBulletsText.defaultColor = newDefaultColor;
                HighlightBulletsBackground.backgroundColor = newBackgroundColor;

            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsPropertiesEdited");
                throw;
            }
        }
        public void HighlightBulletsDialogBoxPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new HighlightBulletsDialogBox(HighlightBulletsText.highlightColor, HighlightBulletsText.defaultColor, HighlightBulletsBackground.backgroundColor);
                dialog.SettingsHandler += HighlightBulletsPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "HighlightBulletsDialogBoxPressed");
                throw;
            }
        }

        public bool GetEmbedAudioVisiblity(Office.IRibbonControl control)
        {
            return EmbedAudioVisible;
        }

        public void RecManagementClick(Office.IRibbonControl control)
        {
            if (!Globals.ThisAddIn.VerifyVersion())
            {
                return;
            }

            Globals.ThisAddIn.RegisterRecorderPane(Globals.ThisAddIn.Application.ActivePresentation);

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));
            var recorder = recorderPane.Control as RecorderTaskPane;

            // TODO:
            // Handle exception when user clicks the button without selecting any slides

            // if currently the pane is hidden, show the pane
            if (recorder != null && !recorderPane.Visible)
            {
                // fire the pane visble change event
                recorderPane.Visible = true;

                // reload the pane
                recorder.RecorderPaneReload();
            }
        }

        # region Custom Shapes
        public void CustomShapeButtonClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.RegisterCustomShapePane(Globals.ThisAddIn.Application.ActivePresentation);
            
            var customShapePane = Globals.ThisAddIn.GetActivePane(typeof(CustomShapePane));

            if (customShapePane == null || !(customShapePane.Control is CustomShapePane))
            {
                return;
            }

            var customShape = customShapePane.Control as CustomShapePane;

            // if currently the pane is hidden, show the pane
            if (customShapePane.Visible)
            {
                return;
            }

            customShape.PaneReload();
            customShapePane.Visible = true;

            Globals.ThisAddIn.InitializeShapeGallery();
        }

        public void AddShapeButtonClick(Office.IRibbonControl control)
        {
            var prensentation = Globals.ThisAddIn.Application.ActivePresentation;
            Globals.ThisAddIn.RegisterCustomShapePane(prensentation);

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            var customShapePane = Globals.ThisAddIn.GetActivePane(typeof(CustomShapePane));

            if (customShapePane == null || !(customShapePane.Control is CustomShapePane))
            {
                return;
            }

            // show pane if not visible
            if (!customShapePane.Visible)
            {
                customShapePane.Visible = true;
            }

            var customShape = customShapePane.Control as CustomShapePane;

            // see below for explanation
            var presentationSaved = prensentation.Saved == Office.MsoTriState.msoTrue &&
                                    prensentation.Path != string.Empty;

            customShape.PaneReload();

            Globals.ThisAddIn.InitializeShapeGallery();

            var shapeName = customShape.NextDefaultNameWithoutExtension;
            var shapeFullName = customShape.NextDefaultFullName;
            // add the selection into pane and save it as .png locally
            ConvertToPicture.ConvertAndSave(selection, shapeFullName);
            customShape.AddCustomShape(shapeName, shapeFullName, true);

            Globals.ThisAddIn.ShapePresentation.AddShape(selection, shapeName);
            Globals.ThisAddIn.ShapePresentation.Save();

            Globals.ThisAddIn.SyncShapeAdd(shapeName, shapeFullName);

            // since we group and then ungroup the shape, document has been modified.
            // if the presentation has been saved before the group->ungroup, we can save
            // the file; else we leave it.
            if (presentationSaved)
            {
                Globals.ThisAddIn.Application.ActivePresentation.Save();
            }
        }
        # endregion

        #region NotesToAudio Button Callbacks
        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }

        public void RemoveAudioClick(Office.IRibbonControl control)
        {
            if (!Globals.ThisAddIn.VerifyVersion())
            {
                return;
            }
            
            try
            {
                NotesToAudio.RemoveAudioFromSelectedSlides();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                throw;
            }

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));
            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null) return;

            recorder.ClearRecordDataListForSelectedSlides();

            // if current list is visible, update the pane immediately
            if (recorderPane.Visible)
            {
                foreach (PowerPointSlide slide in PowerPointCurrentPresentationInfo.SelectedSlides)
                {
                    recorder.UpdateLists(slide.ID);
                }
            }

            RemoveAudioEnabled = false;
            RefreshRibbonControl("RemoveAudioButton");
        }

        public void AddAudioClick(Office.IRibbonControl control)
        {
            if (!Globals.ThisAddIn.VerifyVersion())
            {
                return;
            }

            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

            foreach (PowerPointSlide slide in PowerPointCurrentPresentationInfo.SelectedSlides)
            {
                if (slide.NotesPageText.Trim() != "")
                {
                    RemoveAudioEnabled = true;
                    RefreshRibbonControl("RemoveAudioButton");
                    break;
                }
            }

            var allAudioFiles = NotesToAudio.EmbedSelectedSlideNotes();

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));
            var recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null) return;

            // initialize selected slides' audio
            recorder.InitializeAudioAndScript(PowerPointCurrentPresentationInfo.SelectedSlides.ToList(),
                                                  allAudioFiles, true);
            
            // if current list is visible, update the pane immediately
            if (recorderPane.Visible)
            {
                recorder.UpdateLists(currentSlide.ID);
            }

            PreviewAnimationsIfChecked();
        }

        public void ContextAddAudioClick(Office.IRibbonControl control)
        {
            if (!Globals.ThisAddIn.VerifyVersion())
            {
                return;
            }

            NotesToAudio.EmbedCurrentSlideNotes();
            PreviewAnimationsIfChecked();
        }
        #endregion

        #region NotesToCaptions Button Callbacks

        public void AddCaptionClick(Office.IRibbonControl control)
        {
            foreach (PowerPointSlide slide in PowerPointCurrentPresentationInfo.SelectedSlides)
            {
                if (slide.NotesPageText.Trim() != "")
                {
                    RemoveCaptionsEnabled = true;
                    break;
                }
            }
            NotesToCaptions.EmbedCaptionsOnSelectedSlides();
            RefreshRibbonControl("RemoveCaptionsButton");
        }

        public void RemoveCaptionClick(Office.IRibbonControl control)
        {
            RemoveCaptionsEnabled = false;
            RefreshRibbonControl("RemoveCaptionsButton");
            NotesToCaptions.RemoveCaptionsFromSelectedSlides();
        }

        public void ContextReplaceAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.ReplaceSelectedAudio();
        }

        #endregion

        #region NotesToAudio/Caption Helpers
        public void AutoNarrateDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AutoNarrateDialogBox(_voiceSelected, _voiceNames,
                    _previewCurrentSlide);
                dialog.SettingsHandler += AutoNarrateSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "AutoNarrateDialogButtonPressed");
                throw;
            }
        }

        public void AutoNarrateSettingsChanged(String voiceName, bool previewCurrentSlide)
        {
            _previewCurrentSlide = previewCurrentSlide;
            if (!String.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                _voiceSelected = _voiceNames.IndexOf(voiceName);
            }
        }

        private void PreviewAnimationsIfChecked()
        {
            if (_previewCurrentSlide)
            {
                NotesToAudio.PreviewAnimations();
            }
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

        #region feature: Fit To Slide | Fit To Width | Fit To Height

        public void FitToWidthClick(Office.IRibbonControl control)
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToWidth(selectedShape);
        }

        public void FitToHeightClick(Office.IRibbonControl control)
        {
            var selectedShape = PowerPointCurrentPresentationInfo.CurrentSelection.ShapeRange[1];
            FitToSlide.FitToHeight(selectedShape);
        }

        public Bitmap GetFitToWidthImage(Office.IRibbonControl control)
        {
            return FitToSlide.GetFitToWidthImage(control);
        }

        public Bitmap GetFitToHeightImage(Office.IRibbonControl control)
        {
            return FitToSlide.GetFitToHeightImage(control);
        }

        #endregion

        #region feature: Crop to Shape

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            CropToShape.Crop(selection);
        }

        public Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            return CropToShape.GetCutOutShapeMenuImage(control);
        }

        #endregion

        #region feature: Convert to Picture

        public void ConvertToPictureButtonClick(Office.IRibbonControl control)
        {
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            ConvertToPicture.Convert(selection);
        }

        public Bitmap GetConvertToPicMenuImage(Office.IRibbonControl control)
        {
            return ConvertToPicture.GetConvertToPicMenuImage(control);
        }

        #endregion

        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            const string officeVersion2010 = "14.0";
            return Globals.ThisAddIn.Application.Version == officeVersion2010;
        }

        #region feature: Color
        public void ColorPickerButtonClick(Office.IRibbonControl control)
        {
            try
            {
                ////PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                ////Form ColorPickerForm = new ColorPickerForm(selectedShapes);
                ////ColorPickerForm.Show();
                //ColorDialog MyDialog = new ColorDialog();
                //// Keeps the user from selecting a custom color.
                //MyDialog.AllowFullOpen = false;
                //// Allows the user to get help. (The default is false.)
                //MyDialog.ShowHelp = true;
                //ColorPickerForm colorPickerForm = new ColorPickerForm();
                //colorPickerForm.Show();

                Globals.ThisAddIn.RegisterColorPane(Globals.ThisAddIn.Application.ActivePresentation);

                var colorPane = Globals.ThisAddIn.GetActivePane(typeof(ColorPane));
                var color = colorPane.Control as ColorPane;

                // if currently the pane is hidden, show the pane
                if (!colorPane.Visible)
                {
                    // fire the pane visble change event
                    colorPane.Visible = true;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("No Shape Selected", "Invalid Selection");
                PowerPointLabsGlobals.LogException(e, "ColorPickerButtonClicked");
                throw;
            }
        }
        #endregion

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
    }
}
