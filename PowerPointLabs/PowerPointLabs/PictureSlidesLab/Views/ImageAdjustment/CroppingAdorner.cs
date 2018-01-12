using System;
using System.Diagnostics.CodeAnalysis;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;

namespace PowerPointLabs.PictureSlidesLab.Views.ImageAdjustment
{
    [SuppressMessage("Microsoft.StyleCop.CSharp.OrderingRules", "SA1202:ElementsMustBeOrderedByAccess", Justification = "To refactor to partials")]
    /// <summary>
    /// Taken from
    /// http://www.codeproject.com/Articles/23158/A-Photoshop-like-Cropping-Adorner-for-WPF
    /// Edited to support fixed ratio cropping
    /// </summary>
    public class CroppingAdorner : Adorner
    {
        #region Private variables
        // Width of the thumbs.  I know these really aren't "pixels", but px
        // is still a good mnemonic.
        private const int CpxThumbWidth = 6;

        // PuncturedRect to hold the "Cropping" portion of the adorner
        private PuncturedRect _prCropMask;

        // Canvas to hold the thumbs so they can be moved in response to the user
        private Canvas _cnvThumbs;

        // Cropping adorner uses Thumbs for visual elements.  
        // The Thumbs have built-in mouse input handling.
        private CropThumb _crtTopLeft, _crtTopRight, _crtBottomLeft, _crtBottomRight, _crtCentred;

        // To store and manage the adorner's visual children.
        private VisualCollection _vc;

        public float SlideWidth { get; set; }

        public float SlideHeight { get; set; }

        #endregion

        #region Properties
        public Rect ClippingRectangle
        {
            get
            {
                return _prCropMask.RectInterior;
            }
            set
            {
                _prCropMask.RectInterior = value;
                SetThumbs(_prCropMask.RectInterior);
                RaiseEvent(new RoutedEventArgs(CropChangedEvent, this));
            }
        }
        #endregion

        #region Routed Events
        public static readonly RoutedEvent CropChangedEvent = EventManager.RegisterRoutedEvent(
            "CropChanged",
            RoutingStrategy.Bubble,
            typeof(RoutedEventHandler),
            typeof(CroppingAdorner));

        public event RoutedEventHandler CropChanged
        {
            add
            {
                AddHandler(CroppingAdorner.CropChangedEvent, value);
            }
            remove
            {
                RemoveHandler(CroppingAdorner.CropChangedEvent, value);
            }
        }
        #endregion

        #region Dependency Properties
        static public DependencyProperty FillProperty = Shape.FillProperty.AddOwner(typeof(CroppingAdorner));

        public Brush Fill
        {
            get { return (Brush)GetValue(FillProperty); }
            set { SetValue(FillProperty, value); }
        }

        private static void FillPropChanged(DependencyObject d, DependencyPropertyChangedEventArgs args)
        {
            CroppingAdorner crp = d as CroppingAdorner;

            if (crp != null)
            {
                crp._prCropMask.Fill = (Brush)args.NewValue;
            }
        }
        #endregion

        #region Constructor
        static CroppingAdorner()
        {
            Color clr = Colors.Red;

            clr.A = 80;
            FillProperty.OverrideMetadata(typeof(CroppingAdorner),
                new PropertyMetadata(
                    new SolidColorBrush(clr),
                    new PropertyChangedCallback(FillPropChanged)));
        }

        public CroppingAdorner(UIElement adornedElement, Rect rcInit)
            : base(adornedElement)
        {
            _vc = new VisualCollection(this);
            _prCropMask = new PuncturedRect();
            _prCropMask.IsHitTestVisible = false;
            _prCropMask.RectInterior = rcInit;
            _prCropMask.Fill = Fill;
            _vc.Add(_prCropMask);
            _cnvThumbs = new Canvas();
            _cnvThumbs.HorizontalAlignment = HorizontalAlignment.Stretch;
            _cnvThumbs.VerticalAlignment = VerticalAlignment.Stretch;

            _vc.Add(_cnvThumbs);
            BuildCorner(ref _crtTopLeft, Cursors.SizeNWSE);
            BuildCorner(ref _crtTopRight, Cursors.SizeNESW);
            BuildCorner(ref _crtBottomLeft, Cursors.SizeNESW);
            BuildCorner(ref _crtBottomRight, Cursors.SizeNWSE);
            BuildCorner(ref _crtCentred, Cursors.SizeAll, Brushes.Tomato);

            // Add handlers for Cropping.
            _crtBottomLeft.DragDelta += new DragDeltaEventHandler(HandleBottomLeft);
            _crtBottomRight.DragDelta += new DragDeltaEventHandler(HandleBottomRight);
            _crtTopLeft.DragDelta += new DragDeltaEventHandler(HandleTopLeft);
            _crtTopRight.DragDelta += new DragDeltaEventHandler(HandleTopRight);
            _crtCentred.DragDelta += new DragDeltaEventHandler(HandleCentred);

            // We have to keep the clipping interior withing the bounds of the adorned element
            // so we have to track it's size to guarantee that...
            FrameworkElement fel = adornedElement as FrameworkElement;

            if (fel != null)
            {
                fel.SizeChanged += new SizeChangedEventHandler(AdornedElement_SizeChanged);
            }
        }
        #endregion

        #region Thumb handlers

        // Handler for Cropping from the bottom-left.
        private void HandleBottomLeft(object sender, DragDeltaEventArgs args)
        {
            if (sender is CropThumb)
            {
                Rect rcInterior = _prCropMask.RectInterior;
                Rect rcExterior = _prCropMask.RectExterior;

                var dx = args.HorizontalChange;

                // boundary: when minimized
                if (rcInterior.Width - dx < 0)
                {
                    dx = rcInterior.Width;
                }

                var newLeft = rcInterior.Left + dx;
                var newWidth = rcInterior.Width - dx;
                var newHeight = newWidth / SlideWidth * SlideHeight;

                // boundary: when maximized
                if (newLeft < rcExterior.Left && rcInterior.Top + newHeight > rcExterior.Bottom)
                {
                    newLeft = rcExterior.Left;
                    dx = newLeft - rcInterior.Left;
                    newWidth = rcInterior.Width - dx;
                    newHeight = newWidth / SlideWidth * SlideHeight;
                    if (rcInterior.Top + newHeight > rcExterior.Bottom)
                    {
                        newHeight = rcExterior.Bottom - rcInterior.Top;
                        newWidth = newHeight / SlideHeight * SlideWidth;
                        dx = rcInterior.Width - newWidth;
                        newLeft = rcInterior.Left + dx;
                    }
                }
                else if (newLeft < rcExterior.Left)
                {
                    newLeft = rcExterior.Left;
                    dx = newLeft - rcInterior.Left;
                    newWidth = rcInterior.Width - dx;
                    newHeight = newWidth / SlideWidth * SlideHeight;
                }
                else if (rcInterior.Top + newHeight > rcExterior.Bottom)
                {
                    newHeight = rcExterior.Bottom - rcInterior.Top;
                    newWidth = newHeight / SlideHeight * SlideWidth;
                    dx = rcInterior.Width - newWidth;
                    newLeft = rcInterior.Left + dx;
                }

                rcInterior = new Rect(
                        newLeft,
                        rcInterior.Top,
                        newWidth,
                        newHeight);

                _prCropMask.RectInterior = rcInterior;
                SetThumbs(_prCropMask.RectInterior);
                RaiseEvent(new RoutedEventArgs(CropChangedEvent, this));
            }
        }

        // Handler for Cropping from the bottom-right.
        private void HandleBottomRight(object sender, DragDeltaEventArgs args)
        {
            if (sender is CropThumb)
            {
                var dx = args.HorizontalChange;
                ZoomCroppingRect(dx);
            }
        }

        public void ZoomCroppingRect(double dx)
        {
            Rect rcInterior = _prCropMask.RectInterior;
            Rect rcExterior = _prCropMask.RectExterior;

            // boundary: when minimized
            if (rcInterior.Width + dx < 0)
            {
                dx = -rcInterior.Width;
            }

            var newWidth = rcInterior.Width + dx;
            var newHeight = newWidth / SlideWidth * SlideHeight;

            // boundary: when maximized
            if (rcInterior.Left + newWidth > rcExterior.Right
                && rcInterior.Top + newHeight > rcExterior.Bottom)
            {
                newWidth = rcExterior.Right - rcInterior.Left;
                newHeight = newWidth / SlideWidth * SlideHeight;
                if (rcInterior.Top + newHeight > rcExterior.Bottom)
                {
                    newHeight = rcExterior.Bottom - rcInterior.Top;
                    newWidth = newHeight / SlideHeight * SlideWidth;
                }
            }
            else if (rcInterior.Left + newWidth > rcExterior.Right)
            {
                newWidth = rcExterior.Right - rcInterior.Left;
                newHeight = newWidth / SlideWidth * SlideHeight;
            }
            else if (rcInterior.Top + newHeight > rcExterior.Bottom)
            {
                newHeight = rcExterior.Bottom - rcInterior.Top;
                newWidth = newHeight / SlideHeight * SlideWidth;
            }

            rcInterior = new Rect(
                    rcInterior.Left,
                    rcInterior.Top,
                    newWidth,
                    newHeight);

            _prCropMask.RectInterior = rcInterior;
            SetThumbs(_prCropMask.RectInterior);
            RaiseEvent(new RoutedEventArgs(CropChangedEvent, this));
        }

        // Handler for Cropping from the top-right.
        private void HandleTopRight(object sender, DragDeltaEventArgs args)
        {
            if (sender is CropThumb)
            {
                Rect rcInterior = _prCropMask.RectInterior;
                Rect rcExterior = _prCropMask.RectExterior;

                var dy = args.VerticalChange;

                // boundary: when minimized

                if (rcInterior.Height - dy < 0)
                {
                    dy = rcInterior.Height;
                }

                var newTop = rcInterior.Top + dy;
                var newHeight = rcInterior.Height - dy;
                var newWidth = newHeight / SlideHeight * SlideWidth;

                // boundary: when maximized
                if (newTop < rcExterior.Top && rcInterior.Left + newWidth > rcExterior.Right)
                {
                    newTop = rcExterior.Top;
                    dy = newTop - rcInterior.Top;
                    newHeight = rcInterior.Height - dy;
                    newWidth = newHeight / SlideHeight * SlideWidth;
                    if (rcInterior.Left + newWidth > rcExterior.Right)
                    {
                        newWidth = rcExterior.Right - rcInterior.Left;
                        newHeight = newWidth / SlideWidth * SlideHeight;
                        dy = rcInterior.Height - newHeight;
                        newTop = rcInterior.Top + dy;
                    }
                }
                else if (newTop < rcExterior.Top)
                {
                    newTop = rcExterior.Top;
                    dy = newTop - rcInterior.Top;
                    newHeight = rcInterior.Height - dy;
                    newWidth = newHeight / SlideHeight * SlideWidth;
                }
                else if (rcInterior.Left + newWidth > rcExterior.Right)
                {
                    newWidth = rcExterior.Right - rcInterior.Left;
                    newHeight = newWidth / SlideWidth * SlideHeight;
                    dy = rcInterior.Height - newHeight;
                    newTop = rcInterior.Top + dy;
                }

                rcInterior = new Rect(
                        rcInterior.Left,
                        newTop,
                        newWidth,
                        newHeight);

                _prCropMask.RectInterior = rcInterior;
                SetThumbs(_prCropMask.RectInterior);
                RaiseEvent(new RoutedEventArgs(CropChangedEvent, this));
            }
        }

        // Handler for Cropping from the top-left.
        private void HandleTopLeft(object sender, DragDeltaEventArgs args)
        {
            if (sender is CropThumb)
            {
                Rect rcInterior = _prCropMask.RectInterior;
                Rect rcExterior = _prCropMask.RectExterior;

                var dx = args.HorizontalChange;

                // boundary: when minimized
                if (rcInterior.Width - dx < 0)
                {
                    dx = rcInterior.Width;
                }

                var newWidth = rcInterior.Width - dx;
                var newLeft = rcInterior.Right - newWidth;
                var newHeight = newWidth / SlideWidth * SlideHeight;
                var newTop = rcInterior.Bottom - newHeight;

                // boundary: when maximized
                if (newTop < rcExterior.Top && newLeft < rcExterior.Left)
                {
                    newTop = rcExterior.Top;
                    newHeight = rcInterior.Bottom - newTop;
                    newWidth = newHeight / SlideHeight * SlideWidth;
                    newLeft = rcInterior.Right - newWidth;
                    if (newLeft < rcExterior.Left)
                    {
                        newLeft = rcExterior.Left;
                        newWidth = rcInterior.Right - newLeft;
                        newHeight = newWidth / SlideWidth * SlideHeight;
                        newTop = rcInterior.Bottom - newHeight;
                    }
                }
                else if (newTop < rcExterior.Top)
                {
                    newTop = rcExterior.Top;
                    newHeight = rcInterior.Bottom - newTop;
                    newWidth = newHeight / SlideHeight * SlideWidth;
                    newLeft = rcInterior.Right - newWidth;
                }
                else if (newLeft < rcExterior.Left)
                {
                    newLeft = rcExterior.Left;
                    newWidth = rcInterior.Right - newLeft;
                    newHeight = newWidth / SlideWidth * SlideHeight;
                    newTop = rcInterior.Bottom - newHeight;
                }

                rcInterior = new Rect(
                        newLeft,
                        newTop,
                        newWidth,
                        newHeight);

                _prCropMask.RectInterior = rcInterior;
                SetThumbs(_prCropMask.RectInterior);
                RaiseEvent(new RoutedEventArgs(CropChangedEvent, this));
            }
        }

        private void HandleCentred(object sender, DragDeltaEventArgs args)
        {
            if (sender is CropThumb)
            {
                var dx = args.HorizontalChange;
                var dy = args.VerticalChange;
                MoveCroppingRect(dx, dy);
            }
        }

        public void MoveCroppingRect(double dx, double dy)
        {
            Rect rcInterior = _prCropMask.RectInterior;
            Rect rcExterior = _prCropMask.RectExterior;  

            if (rcInterior.Left + dx < rcExterior.Left)
            {
                dx = rcExterior.Left - rcInterior.Left;
            }
            else if (rcInterior.Left + dx + rcInterior.Width > rcExterior.Right)
            {
                dx = rcExterior.Right - rcInterior.Left - rcInterior.Width;
            }

            if (rcInterior.Top + dy < rcExterior.Top)
            {
                dy = rcExterior.Top - rcInterior.Top;
            }
            else if (rcInterior.Top + dy + rcInterior.Height > rcExterior.Bottom)
            {
                dy = rcExterior.Bottom - rcInterior.Top - rcInterior.Height;
            }

            rcInterior = new Rect(
                    rcInterior.Left + dx,
                    rcInterior.Top + dy,
                    rcInterior.Width,
                    rcInterior.Height);

            _prCropMask.RectInterior = rcInterior;
            SetThumbs(_prCropMask.RectInterior);
            RaiseEvent(new RoutedEventArgs(CropChangedEvent, this));
        }

        #endregion

        #region Other handlers
        private void AdornedElement_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            FrameworkElement fel = sender as FrameworkElement;
            Rect rcInterior = _prCropMask.RectInterior;
            bool fFixupRequired = false;
            double
                intLeft = rcInterior.Left,
                intTop = rcInterior.Top,
                intWidth = rcInterior.Width,
                intHeight = rcInterior.Height;

            if (rcInterior.Left > fel.RenderSize.Width)
            {
                intLeft = fel.RenderSize.Width;
                intWidth = 0;
                fFixupRequired = true;
            }

            if (rcInterior.Top > fel.RenderSize.Height)
            {
                intTop = fel.RenderSize.Height;
                intHeight = 0;
                fFixupRequired = true;
            }

            if (rcInterior.Right > fel.RenderSize.Width)
            {
                intWidth = Math.Max(0, fel.RenderSize.Width - intLeft);
                fFixupRequired = true;
            }

            if (rcInterior.Bottom > fel.RenderSize.Height)
            {
                intHeight = Math.Max(0, fel.RenderSize.Height - intTop);
                fFixupRequired = true;
            }
            if (fFixupRequired)
            {
                _prCropMask.RectInterior = new Rect(intLeft, intTop, intWidth, intHeight);
            }
        }
        #endregion

        #region Arranging/positioning
        private void SetThumbs(Rect rc)
        {
            _crtBottomRight.SetPos(rc.Right, rc.Bottom);
            _crtTopLeft.SetPos(rc.Left, rc.Top);
            _crtTopRight.SetPos(rc.Right, rc.Top);
            _crtBottomLeft.SetPos(rc.Left, rc.Bottom);
            _crtCentred.SetPos(rc.Left + rc.Width / 2, rc.Top + rc.Height / 2);
        }

        // Arrange the Adorners.
        protected override Size ArrangeOverride(Size finalSize)
        {
            Rect rcExterior = new Rect(0, 0, AdornedElement.RenderSize.Width, AdornedElement.RenderSize.Height);
            _prCropMask.RectExterior = rcExterior;
            Rect rcInterior = _prCropMask.RectInterior;
            _prCropMask.Arrange(rcExterior);

            SetThumbs(rcInterior);
            _cnvThumbs.Arrange(rcExterior);
            return finalSize;
        }
        #endregion

        #region Helper functions

        private void BuildCorner(ref CropThumb crt, Cursor crs, Brush color = null)
        {
            if (crt != null)
            {
                return;
            }

            crt = new CropThumb(CpxThumbWidth, color);

            // Set some arbitrary visual characteristics.
            crt.Cursor = crs;

            _cnvThumbs.Children.Add(crt);
        }
        #endregion

        #region Visual tree overrides
        // Override the VisualChildrenCount and GetVisualChild properties to interface with 
        // the adorner's visual collection.
        protected override int VisualChildrenCount { get { return _vc.Count; } }
        protected override Visual GetVisualChild(int index) { return _vc[index]; }
        #endregion

        #region Internal Classes
        class CropThumb : Thumb
        {
            #region Private variables
            int _cpx;
            Brush _color;
            #endregion

            #region Constructor
            internal CropThumb(int cpx, Brush color)
            {
                _cpx = cpx;
                _color = color;
            }
            #endregion

            #region Overrides
            protected override Visual GetVisualChild(int index)
            {
                return null;
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                drawingContext.DrawRoundedRectangle(_color ?? Brushes.White, new Pen(Brushes.Black, 1), new Rect(new Size(_cpx, _cpx)), 1, 1);
            }
            #endregion

            #region Positioning
            internal void SetPos(double x, double y)
            {
                Canvas.SetTop(this, y - _cpx / 2);
                Canvas.SetLeft(this, x - _cpx / 2);
            }
            #endregion
        }
        #endregion
    }

}

