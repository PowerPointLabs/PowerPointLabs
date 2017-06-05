using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using PowerPointLabs.ActionFramework.Common.Factory;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.DataSources;
using PowerPointLabs.DrawingsLab;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.View;
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
        #region Action Framework Factory
        private ActionHandlerFactory ActionHandlerFactory { get; set; }

        private LabelHandlerFactory LabelHandlerFactory { get; set; }

        private ImageHandlerFactory ImageHandlerFactory { get; set; }

        private SupertipHandlerFactory SupertipHandlerFactory { get; set; }

        private ContentHandlerFactory ContentHandlerFactory { get; set; }

        private PressedHandlerFactory PressedHandlerFactory { get; set; }

        private CheckBoxActionHandlerFactory CheckBoxActionHandlerFactory { get; set; }
        #endregion

        #region Action Framework entry point

        public void OnAction(Office.IRibbonControl control)
        {
            var actionHandler = ActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            actionHandler.Execute(control.Id);
        }

        public string GetLabel(Office.IRibbonControl control)
        {
            var labelHandler = LabelHandlerFactory.CreateInstance(control.Id, control.Tag);
            return labelHandler.Get(control.Id);
        }

        public string GetSupertip(Office.IRibbonControl control)
        {
            var supertipHandler = SupertipHandlerFactory.CreateInstance(control.Id, control.Tag);
            return supertipHandler.Get(control.Id);
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            var imageHandler = ImageHandlerFactory.CreateInstance(control.Id, control.Tag);
            return imageHandler.Get(control.Id);
        }

        public string GetContent(Office.IRibbonControl control)
        {
            var contentHandler = ContentHandlerFactory.CreateInstance(control.Id, control.Tag);
            return contentHandler.Get(control.Id);
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            var pressedHandler = PressedHandlerFactory.CreateInstance(control.Id, control.Tag);
            return pressedHandler.Get(control.Id);
        }

        public void OnCheckBoxAction(Office.IRibbonControl control, bool pressed)
        {
            var checkBoxActionHandler = CheckBoxActionHandlerFactory.CreateInstance(control.Id, control.Tag);
            checkBoxActionHandler.Execute(control.Id, pressed);
        }

        #endregion

        #region Deprecated. Please only use Action Framework to support the feature.

#pragma warning disable 0618
        private Office.IRibbonUI _ribbon;
        
        public bool FrameAnimationChecked = false;
        public bool BackgroundZoomChecked = true;
        public bool MultiSlideZoomChecked = true;
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

        public bool HighlightTextFragmentsEnabled = true;

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
            ActionHandlerFactory = new ActionHandlerFactory();
            LabelHandlerFactory = new LabelHandlerFactory();
            SupertipHandlerFactory = new SupertipHandlerFactory();
            ImageHandlerFactory = new ImageHandlerFactory();
            ContentHandlerFactory = new ContentHandlerFactory();
            PressedHandlerFactory = new PressedHandlerFactory();
            CheckBoxActionHandlerFactory = new CheckBoxActionHandlerFactory();

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
                Logger.LogException(e, "RefreshRibbonControl");
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
                Globals.ThisAddIn.Application.StartNewUndoEntry();

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
                Logger.LogException(e, "HighlightBulletsBackgroundButtonClick");
                throw;
            }
        }
        public void HighlightBulletsTextButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

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
                Logger.LogException(e, "HighlightBulletsTextButtonClick");
                throw;
            }
        }
        public void HighlightTextFragmentsButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kShapeSelected;
                else if (Globals.ThisAddIn.Application.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kTextSelected;
                else
                    HighlightTextFragments.userSelection = HighlightTextFragments.HighlightTextSelection.kNoneSelected;

                HighlightTextFragments.AddHighlightedTextFragments();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "HighlightTextFragmentsButtonClick");
                throw;
            }
        }

        public void AddInSlideAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AnimateInSlide.isHighlightBullets = false;
                AnimateInSlide.AddAnimationInSlide();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddInSlideAnimationButtonClick");
                throw;
            }
        }

        public void SpotlightBtnClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                Spotlight.AddSpotlightEffect();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        # region Supertips
        public string GetAddAnimationButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationButtonSupertip;
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

        public string GetSpotlightPropertiesButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.SpotlightPropertiesButtonSupertip;
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
        public string GetRemoveAllNotesButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAllNotesButtonSupertip;
        }
        
        public string GetHighlightBulletsTextButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsTextButtonSupertip;
        }
        public string GetHighlightBulletsBackgroundButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsBackgroundButtonSupertip;
        }

        public string GetHighlightTextFragmentsButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.HighlightTextFragmentsButtonSupertip;
        }
        
        public string GetCustomeShapeButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.CustomeShapeButtonSupertip;
        }

        public string GetEffectsLabSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMenuSupertip;
        }
        public string GetEffectsLabMakeTransparentSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMakeTransparentSupertip;
        }
        public string GetEffectsLabMagnifyGlassSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMagnifyGlassSupertip;
        }
        public string GetEffectsLabBlurRemainderSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurRemainderSupertip;
        }
        public string GetEffectsLabColorizeRemainderSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabColorizeRemainderSupertip;
        }
        public string GetEffectsLabBlurBackgroundSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurBackgroundSupertip;
        }
        public string GetEffectsLabColorizeBackgroundSupertip(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabColorizeBackgroundSupertip;
        }

        public string GetAgendaLabSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabSupertip;
        }
        public string GetAgendaLabBulletPointSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBulletPointSupertip;
        }
        public string GetAgendaLabVisualAgendaSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabVisualAgendaSupertip;
        }
        public string GetAgendaLabBeamAgendaSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBeamAgendaSupertip;
        }
        public string GetAgendaLabUpdateAgendaSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabUpdateAgendaSupertip;
        }
        public string GetAgendaLabRemoveAgendaSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabRemoveAgendaSupertip;
        }
        public string GetAgendaLabAgendaSettingsSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabAgendaSettingsSupertip;
        }
        public string GetAgendaLabBulletAgendaSettingsSupertip(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBulletAgendaSettingsSupertip;
        }

        public string GetDrawingsLabButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.DrawingsLabButtonSupertip;
        }

        public string GetResizeLabButtonSupertip(Office.IRibbonControl control)
        {
            return TextCollection.ResizeLabButtonSupertip;
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
        public string GetPositionsLabSupertip(Office.IRibbonControl control)
        {
            return TextCollection.PositionsLabSupertip;
        }
        # endregion

        # region Button Labels
        public string GetPowerPointLabsAddInsTabLabel(Office.IRibbonControl control)
        {
            return TextCollection.PowerPointLabsAddInsTabLabel;
        }

        public string GetCombineShapesLabel(Office.IRibbonControl control)
        {
            return TextCollection.CombineShapesLabel;
        }

        public string GetAnimationLabGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.AnimationLabGroupLabel;
        }
        public string GetAddAnimationButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationButtonLabel;
        }
        public string GetInSlideAnimateButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddAnimationInSlideAnimateButtonLabel;
        }

        public string GetZoomLabGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomLabGroupLabel;
        }
        public string GetAddZoomInButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomInButtonLabel;
        }
        public string GetAddZoomOutButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddZoomOutButtonLabel;
        }
        public string GetZoomToAreaButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomToAreaButtonLabel;
        }

        public string GetCropLabGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.CropLabGroupLabel;
        }
        public string GetMoveCropShapeButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.MoveCropShapeButtonLabel;
        }
        public string GetAddSpotlightButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddSpotlightButtonLabel;
        }
        public string GetReloadSpotlightButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.ReloadSpotlightButtonLabel;
        }

        public string GetEmbedAudioGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.EmbedAudioGroupLabel;
        }
        public string GetAddAudioButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddAudioButtonLabel;
        }
        public string GetGenerateRecordButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.GenerateRecordButtonLabel;
        }
        public string GetAddRecordButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddRecordButtonLabel;
        }
        public string GetRemoveAudioButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAudioButtonLabel;
        }

        public string GetEmbedCaptionGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.EmbedCaptionGroupLabel;
        }
        public string GetAddCaptionsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddCaptionsButtonLabel;
        }
        public string GetRemoveCaptionsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.RemoveCaptionsButtonLabel;
        }
        public string GetRemoveAllNotesButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.RemoveAllNotesButtonLabel;
        }
        public string GetSpotlightPropertiesButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.SpotlightPropertiesButtonLabel;
        }
        public string GetHighlightLabGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightLabGroupLabel;
        }
        public string GetHighlightBulletsTextButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsTextButtonLabel;
        }
        public string GetHighlightBulletsBackgroundButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsBackgroundButtonLabel;
        }
        public string GetHighlightTextFragmentsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightTextFragmentsButtonLabel;
        }        
        public string GetLabsGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.LabsGroupLabel;
        }
        public string GetCustomeShapeButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.CustomeShapeButtonLabel;
        }

        public string GetEffectsLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabButtonLabel;
        }
        public string GetEffectsLabMakeTransparentButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMakeTransparentButtonLabel;
        }
        public string GetEffectsLabMagnifyGlassButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabMagnifyGlassButtonLabel;
        }
        public string GetEffectsLabBlurRemainderButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurRemainderButtonLabel;
        }
        public string GetEffectsLabBlurBackgroundButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabBlurBackgroundButtonLabel;
        }
        public string GetEffectsLabRecolorRemainderButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabRecolorRemainderButtonLabel;
        }
        public string GetEffectsLabRecolorBackgroundButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.EffectsLabRecolorBackgroundButtonLabel;
        }

        public string GetAgendaLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabButtonLabel;
        }
        public string GetAgendaLabBulletPointButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBulletPointButtonLabel;
        }
        public string GetAgendaLabVisualAgendaButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabVisualAgendaButtonLabel;
        }
        public string GetAgendaLabBeamAgendaButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBeamAgendaButtonLabel;
        }
        public string GetAgendaLabUpdateAgendaButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabUpdateAgendaButtonLabel;
        }
        public string GetAgendaLabRemoveAgendaButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabRemoveAgendaButtonLabel;
        }
        public string GetAgendaLabAgendaSettingsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabAgendaSettingsButtonLabel;
        }
        public string GetAgendaLabBulletAgendaSettingsButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AgendaLabBulletAgendaSettingsButtonLabel;
        }

        public string GetDrawingsLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.DrawingsLabButtonLabel;
        }

        public string GetPositionsLabButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.PositionsLabButtonLabel;
        }

        public string GetPPTLabsHelpGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.PPTLabsHelpGroupLabel;
        }
        public string GetHelpButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.HelpButtonLabel;
        }
        public string GetFeedbackButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.FeedbackButtonLabel;
        }
        public string GetAboutButtonLabel(Office.IRibbonControl control)
        {
            return TextCollection.AboutButtonLabel;
        }

        public string GetNameEditShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.NameEditShapeLabel;
        }
        public string GetSpotlightShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.SpotlightShapeLabel;
        }
        public string GetZoomInContextMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomInContextMenuLabel;
        }
        public string GetZoomOutContextMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomOutContextMenuLabel;
        }
        public string GetZoomToAreaContextMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.ZoomToAreaContextMenuLabel;
        }
        public string GetHighlightBulletsMenuShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsMenuShapeLabel;
        }
        public string GetHighlightBulletsBackgroundShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsBackgroundShapeLabel;
        }
        public string GetHighlightBulletsTextShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.HighlightBulletsTextShapeLabel;
        }
        public string GetConvertToPictureShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.ConvertToPictureShapeLabel;
        }
        public string GetAddCustomShapeShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.AddCustomShapeShapeLabel;
        }
        public string GetHideSelectedShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.HideSelectedShapeLabel;
        }
        public string GetCutOutShapeShapeLabel(Office.IRibbonControl control)
        {
            return TextCollection.CutOutShapeShapeLabel;
        }
        public string GetInSlideAnimateGroupLabel(Office.IRibbonControl control)
        {
            return TextCollection.InSlideAnimateGroupLabel;
        }
        public string GetApplyAutoMotionThumbnailLabel(Office.IRibbonControl control)
        {
            return TextCollection.ApplyAutoMotionThumbnailLabel;
        }
        public string GetContextSpeakSelectedTextLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextSpeakSelectedTextLabel;
        }
        public string GetContextAddCurrentSlideLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextAddCurrentSlideLabel;
        }
        public string GetContextReplaceAudioLabel(Office.IRibbonControl control)
        {
            return TextCollection.ContextReplaceAudioLabel;
        }
        public string GetPowerPointLabsMenuLabel(Office.IRibbonControl control)
        {
            return TextCollection.PowerPointLabsMenuLabel;
        }
        # endregion

        //Button Click Callbacks        
        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AutoAnimate.AddAutoAnimation();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddAnimationButtonClick");
                throw;
            }
        }
        public void ZoomBtnClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

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
                Logger.LogException(e, "HelpButtonClick");
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
                Logger.LogException(e, "FeedbackButtonClick");
                throw;
            }
        }
        public void AddZoomInButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AutoZoom.AddDrillDownAnimation();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddZoomInButtonClick");
                throw;
            }
        }
        public void AddZoomOutButtonClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AutoZoom.AddStepBackAnimation();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AddZoomOutButtonClick");
                throw;
            }
        }

        # region Icon Getters
        public Bitmap GetAddAnimationImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddAnimation);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAddAnimationImage");
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
                Logger.LogException(e, "GetSpotlightImage");
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
                Logger.LogException(e, "GetReloadSpotlightImage");
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
                Logger.LogException(e, "GetHighlightBulletsTextImage");
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
                Logger.LogException(e, "GetHighlightBulletsBackgroundImage");
                throw;
            }
        }

        public Bitmap GetHighlightWordsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HighlightWords);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetHighlightWordsImage");
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
                Logger.LogException(e, "GetHighlightBulletsTextContextImage");
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
                Logger.LogException(e, "GetHighlightBulletsBackgroundContextImage");
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
                Logger.LogException(e, "GetZoomInImage");
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
                Logger.LogException(e, "GetZoomOutImage");
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
                Logger.LogException(e, "GetZoomToAreaImage");
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
                Logger.LogException(e, "GetZoomToAreaContextImage");
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
                Logger.LogException(e, "GetCropShapeImage");
                throw;
            }
        }

        public Bitmap GetShapesLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.ShapesLab);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetShapesLabImage");
                throw;
            }
        }

        public Bitmap GetEffectsLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.EffectsLab);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetEffectsLabImage");
                throw;
            }
        }
        public Bitmap GetMakeTransparentImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.MakeTransparent);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetMakeTransparentImage");
                throw;
            }
        }
        public Bitmap GetMagnifyImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.Magnify);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetMagnifyImage");
                throw;
            }
        }
        public Bitmap GetBlurRemainderImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.BlurRemainder);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetBlurRemainderImage");
                throw;
            }
        }
        public Bitmap GetRecolorRemainderImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.RecolorRemainder);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetRecolorRemainderImage");
                throw;
            }
        }

        public Bitmap GetAgendaLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaLab);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaLabImage");
                throw;
            }
        }
        public Bitmap GetAgendaTextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaText);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaTextImage");
                throw;
            }
        }
        public Bitmap GetAgendaVisualImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaVisual);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaVisualImage");
                throw;
            }
        }
        public Bitmap GetAgendaSidebarImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaSidebar);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaSidebarImage");
                throw;
            }
        }
        public Bitmap GetAgendaRemoveImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaRemove);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaRemoveImage");
                throw;
            }
        }
        public Bitmap GetAgendaSyncImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaSync);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaSyncImage");
                throw;
            }
        }
        public Bitmap GetAgendaSettingsImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AgendaSettings);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetAgendaSettingsImage");
                throw;
            }
        }
        public Bitmap GetDrawingsLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new System.Drawing.Bitmap(Properties.Resources.DrawingLab);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetDrawingsLabImage");
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
                Logger.LogException(e, "GetAboutImage");
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
                Logger.LogException(e, "GetHelpImage");
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
                Logger.LogException(e, "GetFeedbackImage");
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
                Logger.LogException(e, "GetAddAudioImage");
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
                Logger.LogException(e, "GetRemoveAudioImage");
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
                Logger.LogException(e, "GetAddCaptionsImage");
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
                Logger.LogException(e, "GetRemoveCaptionsImage");
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
                Logger.LogException(e, "GetAddAudioContextImage");
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
                Logger.LogException(e, "GetPreviewNarrationContextImage");
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
                Logger.LogException(e, "GetInSlideAnimationImage");
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
                Logger.LogException(e, "GetAddAnimationContextImage");
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
                Logger.LogException(e, "GetAddSpotlightContextImage");
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
                Logger.LogException(e, "GetEditNameContextImage");
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
                Logger.LogException(e, "GetInSlideAnimationContextImage");
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
                Logger.LogException(e, "GetZoomInContextImage");
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
                Logger.LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        public Bitmap GetAddToCustomShapeContextImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.AddToCustomShapes);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetZoomOutContextImage");
                throw;
            }
        }
        public Bitmap GetHideShapeImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.HideShape);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetHideShapeImage");
                throw;
            }
        }
        # endregion

        public void ZoomStyleChanged(Office.IRibbonControl control, bool pressed)
        {
            try
            {
                BackgroundZoomChecked = pressed;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ZoomStyleChanged");
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
                Logger.LogException(e, "ZoomStyleGetPressed");
                throw;
            }
        }

        # region Control Enable
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
        public bool OnGetEnabledHighlightTextFragments(Office.IRibbonControl control)
        {
            return HighlightTextFragmentsEnabled;
        }
        # endregion

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
                Logger.LogException(e, "NameEditBtnClick");
                throw;
            }
        }
        public void ShapeNameEdited(String newName)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                PowerPoint.Shape selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange[1];
                selectedShape.Name = newName;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ShapeNameEdited");
                throw;
            }
        }

        public void AnimationLabDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new AnimationLabDialogBox(DefaultDuration, FrameAnimationChecked);
                dialog.SettingsHandler += AnimationPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "AnimationLabDialogButtonPressed");
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
                Logger.LogException(e, "AnimationPropertiesEdited");
                throw;
            }
        }

        public void ZoomLabDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new ZoomLabDialogBox(BackgroundZoomChecked, MultiSlideZoomChecked);
                dialog.SettingsHandler += ZoomPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "ZoomLabDialogButtonPressed");
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
                Logger.LogException(e, "ZoomPropertiesEdited");
                throw;
            }
        }

        public void SpotlightDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new SpotlightDialogBox(Spotlight.defaultTransparency, Spotlight.defaultSoftEdges,
                    Spotlight.defaultColor);
                dialog.SettingsHandler += SpotlightPropertiesEdited;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SpotlightDialogButtonPressed");
                throw;
            }
        }

        public void SpotlightPropertiesEdited(float newTransparency, float newSoftEdge, Color newColor)
        {
            try
            {
                Spotlight.defaultTransparency = newTransparency;
                Spotlight.defaultSoftEdges = newSoftEdge;
                Spotlight.defaultColor = newColor;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SpotlightPropertiesEdited");
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
                HighlightTextFragments.backgroundColor = newBackgroundColor;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "HighlightBulletsPropertiesEdited");
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
                Logger.LogException(e, "HighlightBulletsDialogBoxPressed");
                throw;
            }
        }

        public bool GetEmbedAudioVisiblity(Office.IRibbonControl control)
        {
            return EmbedAudioVisible;
        }

        private bool IsValidPresentation(PowerPoint.Presentation pres)
        {
            if (!Globals.ThisAddIn.VerifyVersion(pres))
            {
                MessageBox.Show(TextCollection.VersionNotCompatibleErrorMsg);
                return false;
            }

            return true;
        }

        private void PreviewAnimationsIfChecked()
        {
            if (_previewCurrentSlide)
            {
                NotesToAudio.PreviewAnimations();
            }
        }

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

        public Bitmap GetContextMenuImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.PptlabsContextMenu);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetContextMenuImage");
                throw;
            }
        }

        #endregion

        #region Feature: Picture Slides Lab

        public PictureSlidesLabWindow PictureSlidesLabWindow { get; set; }

        public void PictureSlidesLabButtonClick(Office.IRibbonControl control)
        {
            if (PictureSlidesLabWindow == null || !PictureSlidesLabWindow.IsOpen)
            {
                PictureSlidesLabWindow = new PictureSlidesLabWindow();
                PictureSlidesLabWindow.Show();
            }
            else
            {
                PictureSlidesLabWindow.Activate();
            }            
        }

        public Bitmap GetPictureSlidesLabImage(Office.IRibbonControl control)
        {
            try
            {
                return new Bitmap(Properties.Resources.PictureSlidesLab);
            }
            catch (Exception e)
            {
                Logger.LogException(e, "GetPictureSlidesLabImage");
                throw;
            }
        }

        public string GetPictureSlidesLabSupertip(Office.IRibbonControl control)
        {
            return TextCollection.PictureSlidesLabText.PictureSlidesLabSupertip;
        }
        
        #endregion

        #region Feature: Crop to Shape

        public void CropShapeButtonClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            CropToShape.Crop(selection);
        }

        public Bitmap GetCutOutShapeMenuImage(Office.IRibbonControl control)
        {
            return CropToShape.GetCutOutShapeMenuImage(control);
        }

        #endregion

        #region Feature: Convert to Picture

        public void ConvertToPictureButtonClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            ConvertToPicture.Convert(selection);
        }

        public Bitmap GetConvertToPicMenuImage(Office.IRibbonControl control)
        {
            return ConvertToPicture.GetConvertToPicMenuImage(control);
        }

        #endregion

        # region Feature: Combine Shapes
        public bool GetVisibilityForCombineShapes(Office.IRibbonControl control)
        {
            const string officeVersion2010 = "14.0";
            return Globals.ThisAddIn.Application.Version == officeVersion2010;
        }
        # endregion

        # region Feature: Narrations Lab
        public void AddAudioClick(Office.IRibbonControl control)
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

            if (PowerPointCurrentPresentationInfo.SelectedSlides.Any(slide => slide.NotesPageText.Trim() != ""))
            {
                RemoveAudioEnabled = true;
                RefreshRibbonControl("RemoveAudioButton");
            }

            var allAudioFiles = NotesToAudio.EmbedSelectedSlideNotes();

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null) return;

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

        public void NarrationsLabDialogButtonPressed(Office.IRibbonControl control)
        {
            try
            {
                var dialog = new NarrationsLabDialogBox(_voiceSelected, _voiceNames,
                    _previewCurrentSlide);
                dialog.SettingsHandler += NarrationsLabSettingsChanged;
                dialog.ShowDialog();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "NarrationsLabDialogButtonPressed");
                throw;
            }
        }

        public void NarrationsLabSettingsChanged(String voiceName, bool previewCurrentSlide)
        {
            _previewCurrentSlide = previewCurrentSlide;
            if (!String.IsNullOrWhiteSpace(voiceName))
            {
                NotesToAudio.SetDefaultVoice(voiceName);
                _voiceSelected = _voiceNames.IndexOf(voiceName);
            }
        }

        public void ContextAddAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.EmbedCurrentSlideNotes();
            PreviewAnimationsIfChecked();
        }

        public void ContextReplaceAudioClick(Office.IRibbonControl control)
        {
            NotesToAudio.ReplaceSelectedAudio();
        }

        public void RecManagementClick(Office.IRibbonControl control)
        {
            var currentPresentation = PowerPointPresentation.Current.Presentation;

            if (!IsValidPresentation(currentPresentation))
            {
                return;
            }

            // prepare media files
            var tempPath = Globals.ThisAddIn.PrepareTempFolder(currentPresentation);
            Globals.ThisAddIn.PrepareMediaFiles(currentPresentation, tempPath);

            Globals.ThisAddIn.RegisterRecorderPane(currentPresentation.Windows[1], tempPath);

            var recorderPane = Globals.ThisAddIn.GetActivePane(typeof(RecorderTaskPane));
            var recorder = recorderPane.Control as RecorderTaskPane;

            // if currently the pane is hidden, show the pane
            if (recorder != null && !recorderPane.Visible)
            {
                // fire the pane visble change event
                recorderPane.Visible = true;

                // reload the pane
                recorder.RecorderPaneReload();
            }
        }

        public void RemoveAudioClick(Office.IRibbonControl control)
        {
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

            if (recorderPane == null) return;

            var recorder = recorderPane.Control as RecorderTaskPane;
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

        public void SpeakSelectedTextClick(Office.IRibbonControl control)
        {
            NotesToAudio.SpeakSelectedText();
        }
        # endregion

        # region Feature: Captions Lab
        public void AddCaptionClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

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

        public void RemoveAllNotesClick(Office.IRibbonControl control)
        {
            foreach (var slide in PowerPointCurrentPresentationInfo.SelectedSlides)
            {
                slide.NotesPageText = string.Empty;
            }
        }
        # endregion

        # region Feature: Shapes Lab
        public void CustomShapeButtonClick(Office.IRibbonControl control)
        {
            InitCustomShapePane();
        }

        public void AddShapeButtonClick(Office.IRibbonControl control)
        {
            var customShape = InitCustomShapePane();
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            // first of all we check if the shape gallery has been opened correctly
            if (!Globals.ThisAddIn.ShapePresentation.Opened)
            {
                MessageBox.Show(TextCollection.ShapeGalleryInitErrorMsg);
                return;
            }

            // add shape into shape gallery first to reduce flicker
            var shapeName = Globals.ThisAddIn.ShapePresentation.AddShape(selection,
                                                                         TextCollection.CustomShapeDefaultShapeName);

            // add the selection into pane and save it as .png locally
            var shapeFullName = Path.Combine(customShape.CurrentShapeFolderPath, shapeName + ".png");
            ConvertToPicture.ConvertAndSave(selection, shapeFullName);

            // sync the shape among all opening panels
            Globals.ThisAddIn.SyncShapeAdd(shapeName, shapeFullName, customShape.CurrentCategory);

            // finally, add the shape into the panel and waiting for name editing
            customShape.AddCustomShape(shapeName, shapeFullName, true);
        }

        private static CustomShapePane InitCustomShapePane()
        {
            var prensentation = PowerPointPresentation.Current.Presentation;

            Globals.ThisAddIn.InitializeShapesLabConfig();
            Globals.ThisAddIn.InitializeShapeGallery();
            Globals.ThisAddIn.RegisterShapesLabPane(prensentation);

            var customShapePane = Globals.ThisAddIn.GetActivePane(typeof(CustomShapePane));

            if (customShapePane == null || !(customShapePane.Control is CustomShapePane))
            {
                return null;
            }

            var customShape = customShapePane.Control as CustomShapePane;

            Trace.TraceInformation(
                "Before Visible: " +
                string.Format("Pane Width = {0}, Pane Height = {1}, Control Width = {2}, Control Height {3}",
                              customShapePane.Width, customShapePane.Height, customShape.Width, customShape.Height));

            // if currently the pane is hidden, show the pane
            if (!customShapePane.Visible)
            {
                customShapePane.Visible = true;

                customShape.Width = customShapePane.Width - 16;
                customShape.PaneReload();
            }

            return customShape;
        }
        # endregion

        # region Feature: Effects Lab
        public void MagnifyGlassEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;

            PowerPoint.ShapeRange shapeRange;

            try
            {
                shapeRange = selection.ShapeRange;
            }
            catch (Exception)
            {
                MessageBox.Show("Please select an area to magnify.");

                return;
            }

            if (shapeRange.Count > 1 || shapeRange[1].Type == Office.MsoShapeType.msoGroup)
            {
                MessageBox.Show("Only one magnify area is allowed.");
                
                return;
            }

            try
            {
                var croppedShape = CropToShape.Crop(selection, isInPlace: true, handleError: false);

                MagnifyGlassEffect(croppedShape, 1.4f);
            }
            catch (Exception e)
            {
                var errorMessage = CropToShape.GetErrorMessageForErrorCode(e.Message);
                errorMessage = errorMessage.Replace("Crop To Shape", "Magnify");

                MessageBox.Show(errorMessage);
            }
        }

        public void BlurRemainderEffectClick(int percentage)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null) return;

            effectSlide.BlurBackground(percentage, EffectsLab.EffectsLabBlurSelected.IsTintRemainder);
            effectSlide.GetNativeSlide().Select();
        }

        public void GreyScaleRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null) return;

            effectSlide.GreyScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void BlackWhiteRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null) return;

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void GothamRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null) return;

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void SepiaRemainderEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(true);

            if (effectSlide == null) return;

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void BlurBackgroundEffectClick(int percentage)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null) return;

            effectSlide.BlurBackground(percentage, EffectsLab.EffectsLabBlurSelected.IsTintBackground);
            effectSlide.GetNativeSlide().Select();
        }

        public void GreyScaleBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null) return;

            effectSlide.GreyScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void BlackWhiteBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null) return;

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void GothamBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null) return;

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void SepiaBackgroundEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var effectSlide = GenerateEffectSlide(false);

            if (effectSlide == null) return;

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public void TransparentEffectClick(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Application.StartNewUndoEntry();

            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select at least 1 shape");
                return;
            }

            TransparentEffect(selection.ShapeRange);
        }

        private void MagnifyGlassEffect(PowerPoint.Shape shape, float ratio)
        {
            var delta = 0.5f * (ratio - 1);

            shape.Left -= delta * shape.Width;
            shape.Top -= delta * shape.Height;

            shape.Width *= ratio;
            shape.Height *= ratio;

            shape.Shadow.Visible = Office.MsoTriState.msoTrue;
            shape.Shadow.Style = Office.MsoShadowStyle.msoShadowStyleOuterShadow;
            shape.Shadow.Transparency = 0.6f;
            shape.Shadow.Size = 102f;
            shape.Shadow.Blur = 5;
            shape.Shadow.OffsetX = 0;
            shape.Shadow.OffsetY = 2f;

            shape.ThreeD.BevelTopType = Office.MsoBevelType.msoBevelCircle;
            shape.ThreeD.BevelTopInset = 15;
            shape.ThreeD.BevelTopDepth = 3;
            shape.ThreeD.BevelBottomType = Office.MsoBevelType.msoBevelNone;
            shape.ThreeD.PresetLighting = Office.MsoLightRigType.msoLightRigBalanced;
            shape.ThreeD.LightAngle = 145;

            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
        }

        private PowerPointBgEffectSlide GenerateEffectSlide(bool generateOnRemainder)
        {
            var curSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            var selection = PowerPointCurrentPresentationInfo.CurrentSelection;
            PowerPointSlide dupSlide = null;

            try
            {
                var shapeRange = selection.ShapeRange;

                if (shapeRange.Count != 0)
                {
                    dupSlide = curSlide.Duplicate();
                }

                shapeRange.Cut();

                var effectSlide = PowerPointBgEffectSlide.BgEffectFactory(curSlide.GetNativeSlide(), generateOnRemainder);

                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }
                
                PowerPointPresentation.Current.AddAckSlide();

                return effectSlide;
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
            catch (COMException)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                MessageBox.Show("Please select at least 1 shape");
                return null;
            }
            catch (Exception e)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }
                
                ErrorDialogWrapper.ShowDialog("Error", e.Message, e);
                return null;
            }
        }

        private void TransparentEffect(PowerPoint.ShapeRange shapeRange)
        {
            foreach (PowerPoint.Shape shape in shapeRange)
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subShapeRange = shape.Ungroup();
                    TransparentEffect(subShapeRange);
                    subShapeRange.Group();
                }
                else if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                {
                    PlaceholderTransparencyHandler(shape);
                }
                else if (shape.Type == Office.MsoShapeType.msoPicture)
                {
                    PictureTransparencyHandler(shape);
                }
                else if (shape.Type == Office.MsoShapeType.msoLine)
                {
                    LineTransparencyHandler(shape);
                }
                else if (IsTransparentableShape(shape))
                {
                    ShapeTransparencyHandler(shape);
                }
            }
        }

        private bool IsTransparentableShape(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape ||
                   shape.Type == Office.MsoShapeType.msoFreeform;
        }

        private void PictureTransparencyHandler(PowerPoint.Shape picture)
        {
            var rotation = picture.Rotation;

            picture.Rotation = 0;

            var tempPicPath = Path.Combine(Path.GetTempPath(), "tempPic.png");

            Utils.Graphics.ExportShape(picture, tempPicPath);

            var shapeHolder =
                PowerPointCurrentPresentationInfo.CurrentSlide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    picture.Left,
                    picture.Top,
                    picture.Width,
                    picture.Height);

            var oriZOrder = picture.ZOrderPosition;

            picture.Delete();

            // move shape holder to original z-order
            while (shapeHolder.ZOrderPosition > oriZOrder)
            {
                shapeHolder.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }

            shapeHolder.Line.Visible = Office.MsoTriState.msoFalse;
            shapeHolder.Fill.UserPicture(tempPicPath);
            shapeHolder.Fill.Transparency = 0.5f;

            shapeHolder.Rotation = rotation;

            File.Delete(tempPicPath);
        }

        private void PlaceholderTransparencyHandler(PowerPoint.Shape picture)
        {
            PictureTransparencyHandler(picture);
        }

        private void LineTransparencyHandler(PowerPoint.Shape shape)
        {
            shape.Line.Transparency = 0.5f;
        }

        private void ShapeTransparencyHandler(PowerPoint.Shape shape)
        {
            shape.Fill.Transparency = 0.5f;
            shape.Line.Transparency = 0.5f;
        }
        # endregion

        # region Feature: Agenda Lab
        public void BeamAgendaClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AgendaLab.AgendaLabMain.GenerateAgenda(AgendaLab.Type.Beam);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Unexpected error during generation", e.Message, e);
                throw e;
            }
            GC.Collect();
        }

        public void BulletPointAgendaClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AgendaLab.AgendaLabMain.GenerateAgenda(AgendaLab.Type.Bullet);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Unexpected error during generation", e.Message, e);
                throw e;
            }
            GC.Collect();
        }

        public void VisualAgendaClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AgendaLab.AgendaLabMain.GenerateAgenda(AgendaLab.Type.Visual);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Unexpected error during generation", e.Message, e);
                throw e;
            }
            GC.Collect();
        }

        public void RemoveAgendaClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AgendaLab.AgendaLabMain.RemoveAgenda();
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Unexpected error during removal", e.Message, e);
                throw e;
            }
            GC.Collect();
        }

        public void UpdateAgendaClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.Application.StartNewUndoEntry();

                AgendaLab.AgendaLabMain.SynchroniseAgenda();
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Unexpected error during sync", e.Message, e);
                throw e;
            }
            GC.Collect();
        }
        # endregion

        #region Feature: Drawing Lab
        internal DrawingLabData DrawingLabData { get; set; }
        internal DrawingsLabMain DrawingLab { get; set; }

        public void DrawingsLabButtonClick(Office.IRibbonControl control)
        {
            try
            {
                if (DrawingLabData == null)
                {
                    DrawingLabData = new DrawingLabData();
                    DrawingLab = new DrawingsLabMain(DrawingLabData);
                }

                Globals.ThisAddIn.RegisterDrawingsPane(PowerPointPresentation.Current.Presentation);

                var drawingsPane = Globals.ThisAddIn.GetActivePane(typeof(DrawingsPane));
                ((DrawingsPane)drawingsPane.Control).drawingsPaneWPF.TryInitialise(DrawingLabData, DrawingLab);
                
                // if currently the pane is hidden, show the pane
                if (!drawingsPane.Visible)
                {
                    // fire the pane visble change event
                    drawingsPane.Visible = true;
                }
                else
                {
                    drawingsPane.Visible = false;
                }
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Error in drawing lab", e.Message, e);
                Logger.LogException(e, "DrawingsLabButtonClicked");
                throw;
            }
        }
        #endregion

        // TODO: Add the image for the icon on the ribbon bar
        //public Bitmap GetPositionsLabImage(Office.IRibbonControl control)
        //{
        //    try
        //    {
        //        return new Bitmap(Properties.Resources.PositionsLab);
        //    }
        //    catch (Exception e)
        //    {
        //        PowerPointLabsGlobals.LogException(e, "GetPositionsLabImage");
        //        throw;
        //    }
        //}

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

        public void HideShapeButtonClick(Office.IRibbonControl control)
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            selectedShapes.Visible = Office.MsoTriState.msoFalse;
        }

        #endregion
    }
}
