using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Models;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for GoToSlideDialog.xaml
    /// </summary>
    public partial class GoToSlideDialog
    {
        private const int PreviewHeight = 300;

        public delegate void GotoSlideEvent();

        public delegate void CancelEvent();

        public event GotoSlideEvent OnGotoSlide;

        public event CancelEvent OnCancel;

        public ObservableCollection<ImageItem> SlideList { get; set; }

        public ObservableString DialogTitleProperty { get; set; }

        private int _prevSlideIndex = -1;

        private int _nextSlideIndex = -1;

        public int SelectedSlide { get; set; }

        public GoToSlideDialog()
        {
            InitializeComponent();
            SlideList = new ObservableCollection<ImageItem>();
            DialogTitleProperty = new ObservableString();
            SlideListBox.DataContext = this;
            DialogTitle.DataContext = DialogTitleProperty;
        }

        public void FocusOkButton()
        {
            CancelButton.Focusable = true;
            CancelButton.Focus();
        }

        /// <summary>
        /// This method can only be called after dialog is fully initialized
        /// </summary>
        /// <param name="title"></param>
        public void Init(string title)
        {
            DialogTitleProperty.Text = title;
            Dispatcher.Invoke(new Action(() =>
            {
                SlideList.Clear();
            }));
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
            if (currentSlide.Index - 2 >= 0)
            {
                _prevSlideIndex = currentSlide.Index - 2;
                var prevSlide = PowerPointPresentation.Current.Slides[currentSlide.Index - 2];
                AddSlideThumbnail(prevSlide);
            }
            else
            {
                _prevSlideIndex = currentSlide.Index - 1;
            }

            AddSlideThumbnail(currentSlide, isCurrentSlide: true);

            if (currentSlide.Index < PowerPointPresentation.Current.SlideCount)
            {
                _nextSlideIndex = currentSlide.Index;
                var nextSlide = PowerPointPresentation.Current.Slides[currentSlide.Index];
                AddSlideThumbnail(nextSlide);
            }
            else
            {
                _nextSlideIndex = currentSlide.Index - 1;
            }

            if (SlideListBox.Items.Count == 3)
            {
                SlideListBox.SelectedIndex = 1;
            }
            else if (SlideListBox.Items.Count == 2)
            {
                SlideListBox.SelectedIndex = 1;
            }
            else
            {
                SlideListBox.SelectedIndex = 0;
            }
            SlideListBox.ScrollIntoView(SlideListBox.Items[SlideListBox.SelectedIndex]);
        }

        private void AddSlideThumbnail(PowerPointSlide slide, int pos = -1, bool isCurrentSlide = false)
        {
            if (slide == null) return;

            var thumbnailPath = TempPath.GetPath("slide-" + DateTime.Now.GetHashCode() + slide.Index);
            slide.GetNativeSlide().Export(thumbnailPath, "JPG", GetPreviewWidth(), PreviewHeight);

            ImageItem imageItem;
            if (isCurrentSlide)
            {
                imageItem = new ImageItem
                {
                    ImageFile = thumbnailPath,
                    Tooltip = "(Current) Slide " + slide.Index
                };
            }
            else
            {
                imageItem = new ImageItem
                {
                    ImageFile = thumbnailPath,
                    Tooltip = "Slide " + slide.Index
                };
            }
            
            Dispatcher.Invoke(new Action(() =>
            {
                if (pos == -1)
                {
                    SlideList.Add(imageItem);
                }
                else
                {
                    SlideList.Insert(pos, imageItem);
                }
            }));
        }

        private int GetPreviewWidth()
        {
            return (int)(PowerPointPresentation.Current.SlideWidth / PowerPointPresentation.Current.SlideHeight * PreviewHeight);
        }

        private void GotoSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (OnGotoSlide != null)
            {
                OnGotoSlide();
            }
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (OnCancel != null)
            {
                OnCancel();
            }
        }

        private void SlideListBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SlideListBox.SelectedItem == null)
            {
                GotoSlideButton.IsEnabled = false;
            }
            else
            {
                GotoSlideButton.IsEnabled = true;
                // obtain selected slide
                var selectedItem = (ImageItem) SlideListBox.SelectedItem;
                if (selectedItem.Tooltip.Contains("Current"))
                {
                    SelectedSlide = Int32.Parse(selectedItem.Tooltip.Substring(16));   
                }
                else
                {
                    SelectedSlide = Int32.Parse(selectedItem.Tooltip.Substring(6));
                }

                // auto-load slides when using arrow-keys to navigate
                if (SlideListBox.SelectedIndex == 0)
                {
                    LoadPreviousSlides(isToSelectPrevSlide: false);
                }
                else if (SlideListBox.SelectedIndex == SlideListBox.Items.Count - 1)
                {
                    LoadNextSlides(isToSelectNextSlide: false);
                }
            }
        }

        private void PrevButton_OnClick(object sender, RoutedEventArgs e)
        {
            var selectedIndex = SlideListBox.SelectedIndex;
            if (selectedIndex - 1 >= 0)
            {
                SlideListBox.SelectedIndex = selectedIndex - 1;
                SlideListBox.ScrollIntoView(SlideListBox.Items[SlideListBox.SelectedIndex]);
            }
            else
            {
                LoadPreviousSlides();
            }
        }

        private void LoadPreviousSlides(bool isToSelectPrevSlide = true)
        {
            var newPrevSlideIndex = _prevSlideIndex;

            if (_prevSlideIndex - 3 >= 0)
            {
                newPrevSlideIndex = _prevSlideIndex - 3;
                var prevSlide = PowerPointPresentation.Current.Slides[_prevSlideIndex - 3];
                AddSlideThumbnail(prevSlide, 0);
            }

            if (_prevSlideIndex - 2 >= 0)
            {
                var prevSlide = PowerPointPresentation.Current.Slides[_prevSlideIndex - 2];
                if (_prevSlideIndex - 3 >= 0)
                {
                    AddSlideThumbnail(prevSlide, 1);
                }
                else
                {
                    newPrevSlideIndex = _prevSlideIndex - 2;
                    AddSlideThumbnail(prevSlide, 0);
                }
            }

            if (_prevSlideIndex - 1 >= 0)
            {
                var prevSlide = PowerPointPresentation.Current.Slides[_prevSlideIndex - 1];
                if (_prevSlideIndex - 3 >= 0)
                {
                    AddSlideThumbnail(prevSlide, 2);
                }
                else if (_prevSlideIndex - 2 >= 0)
                {
                    AddSlideThumbnail(prevSlide, 1);
                }
                else
                {
                    newPrevSlideIndex = _prevSlideIndex - 1;
                    AddSlideThumbnail(prevSlide, 0);
                }
            }

            _prevSlideIndex = newPrevSlideIndex;
            if (SlideListBox.SelectedIndex - 1 >= 0 && isToSelectPrevSlide)
            {
                SlideListBox.SelectedIndex = SlideListBox.SelectedIndex - 1;
            }
            SlideListBox.ScrollIntoView(SlideListBox.Items[SlideListBox.SelectedIndex]);
        }

        private void NextButton_OnClick(object sender, RoutedEventArgs e)
        {
            var selectedIndex = SlideListBox.SelectedIndex;
            if (selectedIndex + 1 < SlideListBox.Items.Count)
            {
                SlideListBox.SelectedIndex = selectedIndex + 1;
                SlideListBox.ScrollIntoView(SlideListBox.Items[SlideListBox.SelectedIndex]);
            }
            else
            {
                LoadNextSlides();
            }
        }

        private void LoadNextSlides(bool isToSelectNextSlide = true)
        {
            var newNextSlideIndex = _nextSlideIndex;

            if (_nextSlideIndex + 1 < PowerPointPresentation.Current.SlideCount)
            {
                var nextSlide = PowerPointPresentation.Current.Slides[_nextSlideIndex + 1];
                newNextSlideIndex = _nextSlideIndex + 1;
                AddSlideThumbnail(nextSlide);
            }

            if (_nextSlideIndex + 2 < PowerPointPresentation.Current.SlideCount)
            {
                var nextSlide = PowerPointPresentation.Current.Slides[_nextSlideIndex + 2];
                newNextSlideIndex = _nextSlideIndex + 2;
                AddSlideThumbnail(nextSlide);
            }

            if (_nextSlideIndex + 3 < PowerPointPresentation.Current.SlideCount)
            {
                var nextSlide = PowerPointPresentation.Current.Slides[_nextSlideIndex + 3];
                newNextSlideIndex = _nextSlideIndex + 3;
                AddSlideThumbnail(nextSlide);
            }

            _nextSlideIndex = newNextSlideIndex;
            if (SlideListBox.SelectedIndex + 1 < SlideListBox.Items.Count && isToSelectNextSlide)
            {
                SlideListBox.SelectedIndex = SlideListBox.SelectedIndex + 1;
            }
            SlideListBox.ScrollIntoView(SlideListBox.Items[SlideListBox.SelectedIndex]);
        }
    }
}
