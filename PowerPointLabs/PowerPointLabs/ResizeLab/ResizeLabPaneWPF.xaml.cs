using System;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    /// <summary>
    /// Interaction logic for ResizePane.xaml
    /// </summary>
    public partial class ResizeLabPaneWPF : IResizeLabPane
    {
        private ResizeLabMain _resizeLab;
        private readonly ResizeLabErrorHandler _errorHandler;
        public static bool IsAspectRatioLocked { get; set; }
        
        private const string UnlockAspectRatioToolTip = "Unlocks the aspect ratio of objects when performing resizing of objects";
        private const string LockAspectRatioToolTip = "Locks the aspect ratio of objects when performing resizing of objects";

        // Dialog windows
        private StretchSettingsDialog _stretchSettingsDialog;
        private SameDimensionSettingsDialog _sameDimensionSettingsDialog;

        // For preview
        private Thread thread;
        private const int PreviewDelay = 400;

        public ResizeLabPaneWPF()
        {
            InitializeComponent();
            InitialiseLogicInstance();
            _errorHandler = ResizeLabErrorHandler.InitializErrorHandler(this);
            UnlockAspectRatio();
        }

        internal void InitialiseLogicInstance()
        {
            if (_resizeLab == null)
            {
                _resizeLab = new ResizeLabMain();
            }
        }

        #region Execute Stretch and Shrink

        private void StretchLeftBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.StretchLeft(shapes);
            ClickHandler(resizeAction, ResizeLabMain.Stretch_MinNoOfShapesRequired,
                ResizeLabMain.Stretch_ErrorParameters);
        }

        private void StretchRightBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.StretchRight(shapes);
            ClickHandler(resizeAction, ResizeLabMain.Stretch_MinNoOfShapesRequired,
                ResizeLabMain.Stretch_ErrorParameters);
        }

        private void StretchTopBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.StretchTop(shapes);
            ClickHandler(resizeAction, ResizeLabMain.Stretch_MinNoOfShapesRequired,
                ResizeLabMain.Stretch_ErrorParameters);
        }

        private void StretchBottomBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.StretchBottom(shapes);
            ClickHandler(resizeAction, ResizeLabMain.Stretch_MinNoOfShapesRequired,
                ResizeLabMain.Stretch_ErrorParameters);
        }

        private void StretchSettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_stretchSettingsDialog == null || !_stretchSettingsDialog.IsOpen)
            {
                _stretchSettingsDialog = new StretchSettingsDialog(_resizeLab);
                _stretchSettingsDialog.Show();
            }
            else
            {
                _stretchSettingsDialog.Activate();
            }
            
        }

        #endregion

        #region Execute Same Dimension

        private void SameWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.ResizeToSameWidth(shapes);
            ClickHandler(resizeAction, ResizeLabMain.SameDimension_MinNoOfShapesRequired,
                ResizeLabMain.SameDimension_ErrorParameters);
        }

        private void SameHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.ResizeToSameHeight(shapes);
            ClickHandler(resizeAction, ResizeLabMain.SameDimension_MinNoOfShapesRequired,
                            ResizeLabMain.SameDimension_ErrorParameters);
        }

        private void SameSizeBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.ResizeToSameHeightAndWidth(shapes);
            ClickHandler(resizeAction, ResizeLabMain.SameDimension_MinNoOfShapesRequired,
                ResizeLabMain.SameDimension_ErrorParameters);
        }

        private void SameDimensionSettingsBtn_Click(object sender, RoutedEventArgs e)
        {
            if (_sameDimensionSettingsDialog == null || !_sameDimensionSettingsDialog.IsOpen)
            {
                _sameDimensionSettingsDialog = new SameDimensionSettingsDialog(_resizeLab);
                _sameDimensionSettingsDialog.Show();
            }
            else
            {
                _sameDimensionSettingsDialog.Activate();
            }
        }

        #endregion

        #region Execute Fit
        private void FitWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange, float, float, bool> resizeAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToWidth(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };
            ClickHandler(resizeAction, ResizeLabMain.Fit_MinNoOfShapesRequired,
                ResizeLabMain.Fit_ErrorParameters);
        }

        private void FitHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange, float, float, bool> resizeAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToHeight(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };
            ClickHandler(resizeAction, ResizeLabMain.Fit_MinNoOfShapesRequired,
                            ResizeLabMain.Fit_ErrorParameters);
        }

        private void FillBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange, float, float, bool> resizeAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToFill(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };
            ClickHandler(resizeAction, ResizeLabMain.Fit_MinNoOfShapesRequired,
                ResizeLabMain.Fit_ErrorParameters);
        }

        #endregion

        #region Execute Slight Adjust
        private void IncreaseHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.IncreaseHeight(shapes);
            ClickHandler(resizeAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired,
                ResizeLabMain.SlightAdjust_ErrorParameters);
        }

        private void DecreaseHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.DecreaseHeight(shapes);
            ClickHandler(resizeAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired,
                ResizeLabMain.SlightAdjust_ErrorParameters);
        }

        private void IncreaseWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.IncreaseWidth(shapes);
            ClickHandler(resizeAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired,
                ResizeLabMain.SlightAdjust_ErrorParameters);
        }

        private void DecreaseWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.DecreaseWidth(shapes);
            ClickHandler(resizeAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired,
                ResizeLabMain.SlightAdjust_ErrorParameters);
        }

        #endregion

        #region Execute Aspect Ratio

        private void LockAspectRatio_UnChecked(object sender, RoutedEventArgs e)
        {
            UnlockAspectRatio();
        }

        private void LockAspectRatio_Checked(object sender, RoutedEventArgs e)
        {
            LockAspectRatio();
        }

        private void UnlockAspectRatio()
        {
            IsAspectRatioLocked = false;
            LockAspectRatioCheckBox.ToolTip = LockAspectRatioToolTip;

            ModifySelectionAspectRatio();
        }

        private void LockAspectRatio()
        {
            IsAspectRatioLocked = true;
            LockAspectRatioCheckBox.ToolTip = UnlockAspectRatioToolTip;

            ModifySelectionAspectRatio();
        }

        private void ModifySelectionAspectRatio()
        {
            if (_resizeLab.IsSelectionValid(GetSelection(), false))
            {
                _resizeLab.ChangeShapesAspectRatio(GetSelectedShapes(), IsAspectRatioLocked);
            }
        }

        #endregion

        #region Execute Adjust Aspect Ratio
        private void MatchWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.MatchWidth(shapes);
            ClickHandler(resizeAction, ResizeLabMain.Match_MinNoOfShapesRequired,
                ResizeLabMain.Match_ErrorParameters);
        }

        private void MatchHeightBtn_Click(object sender, RoutedEventArgs e)
        {
            Action<PowerPoint.ShapeRange> resizeAction = shapes => _resizeLab.MatchHeight(shapes);
            ClickHandler(resizeAction, ResizeLabMain.Match_MinNoOfShapesRequired,
                ResizeLabMain.Match_ErrorParameters);
        }

        #endregion

        #region Preview Stretch and Shrink

        private void StretchLeftBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.StretchLeft(shapes);
            PreviewHandler(previewAction, ResizeLabMain.Stretch_MinNoOfShapesRequired);
        }

        private void StretchRightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.StretchRight(shapes);
            PreviewHandler(previewAction, ResizeLabMain.Stretch_MinNoOfShapesRequired);
        }

        private void StretchTopBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.StretchTop(shapes);
            PreviewHandler(previewAction, ResizeLabMain.Stretch_MinNoOfShapesRequired);
        }

        private void StretchBottomBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.StretchBottom(shapes);
            PreviewHandler(previewAction, ResizeLabMain.Stretch_MinNoOfShapesRequired);
        }

        #endregion

        #region Preview Same Dimension

        private void SameWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.ResizeToSameWidth(shapes);
            PreviewHandler(previewAction, ResizeLabMain.SameDimension_MinNoOfShapesRequired);
        }
        
        private void SameHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.ResizeToSameHeight(shapes);
            PreviewHandler(previewAction, ResizeLabMain.SameDimension_MinNoOfShapesRequired);
        }

        private void SameSizeBtn_MouseEnter(object sender, MouseEventArgs e)
        { 
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.ResizeToSameHeightAndWidth(shapes);

            PreviewHandler(previewAction, ResizeLabMain.SameDimension_MinNoOfShapesRequired);
        }

        #endregion

        #region Preview Fit

        private void FitWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange, float, float, bool> previewAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToWidth(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };

            PreviewHandler(previewAction);
        }

        private void FitHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange, float, float, bool> previewAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToHeight(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };
            PreviewHandler(previewAction);
        }

        private void FillBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange, float, float, bool> previewAction =
                (shapes, referenceWidth, referenceHeight, isAspectRatio) =>
                {
                    _resizeLab.FitToFill(shapes, referenceWidth, referenceHeight, isAspectRatio);
                };
            PreviewHandler(previewAction);
        }

        #endregion

        #region Preview Slight Adjust
        private void IncreaseHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.IncreaseHeight(shapes);
            PreviewHandler(previewAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired);
        }

        private void DecreaseHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.DecreaseHeight(shapes);
            PreviewHandler(previewAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired);
        }

        private void IncreaseWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.IncreaseWidth(shapes);
            PreviewHandler(previewAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired);
        }

        private void DecreaseWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.DecreaseWidth(shapes);
            PreviewHandler(previewAction, ResizeLabMain.SlightAdjust_MinNoOfShapesRequired);
        }

        #endregion

        #region Preview Adjust Aspect Ratio
        private void MatchWidthBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.MatchWidth(shapes);
            PreviewHandler(previewAction, ResizeLabMain.Match_MinNoOfShapesRequired);
        }

        private void MatchHeightBtn_MouseEnter(object sender, MouseEventArgs e)
        {
            Action<PowerPoint.ShapeRange> previewAction = shapes => _resizeLab.MatchHeight(shapes);
            PreviewHandler(previewAction, ResizeLabMain.Match_MinNoOfShapesRequired);
        }

        #endregion

        #region Miscellaneous events
        private void Btn_MouseLeave(object sender, MouseEventArgs e)
        {
            if (thread != null && thread.IsAlive) // Actual preview did not execute
            {
                thread.Abort();
            }
            else // Preview was executed
            {
                Reset();
            }
            thread = null;
        }

        #endregion

        #region Helper Functions

        private PowerPoint.ShapeRange GetSelectedShapes(bool handleError = false)
        {
            var selection = GetSelection();

            return _resizeLab.IsSelectionValid(selection, handleError) ? GetSelection().ShapeRange : null;
        }

        private PowerPoint.Selection GetSelection()
        {
            return this.GetCurrentSelection();
        }

        private void ClickHandler(Action<PowerPoint.ShapeRange> resizeAction, int minNoOfSelectedShapes, string[] errorParameters)
        {
            var selectedShapes = GetSelectedShapes();

            if (selectedShapes == null || selectedShapes.Count < minNoOfSelectedShapes)
            {
                _errorHandler.ProcessErrorCode(ResizeLabErrorHandler.ErrorCodeInvalidSelection, errorParameters);
                return;
            }

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, resizeAction);
        }

        private void ClickHandler(Action<PowerPoint.ShapeRange, float, float, bool> resizeAction, int minNoOfSelectedShapes,
            string[] errorParameters)
        {
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;

            if (selectedShapes == null || selectedShapes.Count < minNoOfSelectedShapes)
            {
                _errorHandler.ProcessErrorCode(ResizeLabErrorHandler.ErrorCodeInvalidSelection, errorParameters);
                return;
            }

            ModifySelectionAspectRatio();
            ExecuteResizeAction(selectedShapes, slideWidth, slideHeight, resizeAction);
        }

        private void PreviewHandler(Action<PowerPoint.ShapeRange> previewAction, int minNoOfSelectedShapes)
        {
            thread = new Thread(() => PreviewHandlerAction(previewAction, minNoOfSelectedShapes));
            thread.Start(); 
        }

        private void PreviewHandler(Action<PowerPoint.ShapeRange, float, float, bool> previewAction)
        {
            thread = new Thread(() => PreviewHandlerAction(previewAction));
            thread.Start();
        }

        private void PreviewHandlerAction(Action<PowerPoint.ShapeRange> previewAction, int minNoOfSelectedShapes)
        {
            Thread.Sleep(PreviewDelay);
            var selectedShapes = GetSelectedShapes();

            ModifySelectionAspectRatio();
            Preview(selectedShapes, previewAction, minNoOfSelectedShapes);
        }

        private void PreviewHandlerAction(Action<PowerPoint.ShapeRange, float, float, bool> previewAction)
        {
            Thread.Sleep(PreviewDelay);
            var selectedShapes = GetSelectedShapes();
            var slideWidth = this.GetCurrentPresentation().SlideWidth;
            var slideHeight = this.GetCurrentPresentation().SlideHeight;

            ModifySelectionAspectRatio();
            Preview(selectedShapes, slideWidth, slideHeight, previewAction);
        }


        #endregion
    }
}
