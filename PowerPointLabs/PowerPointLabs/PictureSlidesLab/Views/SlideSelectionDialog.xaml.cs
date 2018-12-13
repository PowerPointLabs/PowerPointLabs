using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    /// <summary>
    /// Interaction logic for SlideSelectionDialog.xaml
    /// </summary>
    public partial class SlideSelectionDialog
    {
        private const int PreviewHeight = 300;

        public delegate void OkEvent();

        public delegate void CancelEvent();

        public event OkEvent OnGotoSlide;

        public event CancelEvent OnCancel;

        public event OkEvent OnAdditionalButtonClick;

        public ObservableCollection<ImageItem> SlideList { get; set; }

        public ObservableString DialogTitleProperty { get; set; }

        public bool IsOpen { get; set; }

        private int _prevSlideIndex = -1;

        private int _nextSlideIndex = -1;

        public int SelectedSlide { get; set; }

        public SlideSelectionDialog()
        {
            InitializeComponent();
            SlideList = new ObservableCollection<ImageItem>();
            DialogTitleProperty = new ObservableString();
            SlideListBox.DataContext = this;
            DialogTitle.DataContext = DialogTitleProperty;
        }

        public SlideSelectionDialog FocusOkButton()
        {
            CancelButton.Focusable = true;
            CancelButton.Focus();
            return this;
        }

        /// <summary>
        /// This method can only be called after dialog is fully initialized
        /// </summary>
        /// <param name="title"></param>
        public SlideSelectionDialog Init(string title)
        {
            DialogTitleProperty.Text = title;
            Dispatcher.Invoke(new Action(() =>
            {
                SlideList.Clear();
            }));
            PowerPointSlide currentSlide = this.GetCurrentSlide();
            if (currentSlide.Index - 2 >= 0)
            {
                _prevSlideIndex = currentSlide.Index - 2;
                PowerPointSlide prevSlide = this.GetCurrentPresentation().Slides[currentSlide.Index - 2];
                AddSlideThumbnail(prevSlide);
            }
            else
            {
                _prevSlideIndex = currentSlide.Index - 1;
            }

            AddSlideThumbnail(currentSlide, isCurrentSlide: true);

            if (currentSlide.Index < this.GetCurrentPresentation().SlideCount)
            {
                _nextSlideIndex = currentSlide.Index;
                PowerPointSlide nextSlide = this.GetCurrentPresentation().Slides[currentSlide.Index];
                AddSlideThumbnail(nextSlide);
            }
            else
            {
                _nextSlideIndex = currentSlide.Index - 1;
            }

            SelectCurrentSlide();

            LoadNextSlides(false);
            SlideListBox.ScrollIntoView(SlideListBox.SelectedItem);
            return this;
        }

        public void OpenDialog()
        {
            IsOpen = true;
        }

        public void CloseDialog()
        {
            IsOpen = false;
        }


        public SlideSelectionDialog CustomizeGotoSlideButton(string content, string tooltip)
        {
            GotoSlideButton.Content = content;
            GotoSlideButton.ToolTip = tooltip;
            return this;
        }

        public SlideSelectionDialog CustomizeAdditionalButton(string content, string tooltip)
        {
            AdditionalButton.Content = content;
            AdditionalButton.ToolTip = tooltip;
            AdditionalButton.Visibility = Visibility.Visible;
            return this;
        }

        private void SelectCurrentSlide()
        {
            foreach (ImageItem slide in SlideList)
            {
                if (slide.Tooltip.Contains("Current"))
                {
                    SlideListBox.SelectedItem = slide;
                }
            }
        }

        private void AddSlideThumbnail(PowerPointSlide slide, int pos = -1, bool isCurrentSlide = false)
        {
            if (slide == null)
            {
                return;
            }

            string thumbnailPath = TempPath.GetPath("slide-" + DateTime.Now.GetHashCode() + slide.Index);
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
            return (int)(this.GetCurrentPresentation().SlideWidth / this.GetCurrentPresentation().SlideHeight * PreviewHeight);
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

        private void AdditionalButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (OnAdditionalButtonClick != null)
            {
                OnAdditionalButtonClick();
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
                ImageItem selectedItem = (ImageItem) SlideListBox.SelectedItem;
                if (selectedItem.Tooltip.Contains("Current"))
                {
                    SelectedSlide = Int32.Parse(selectedItem.Tooltip.Substring(16));   
                }
                else
                {
                    SelectedSlide = Int32.Parse(selectedItem.Tooltip.Substring(6));
                }
            }
        }

        private void PrevButton_OnClick(object sender, RoutedEventArgs e)
        {
            int selectedIndex = SlideListBox.SelectedIndex;
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
            int newPrevSlideIndex = _prevSlideIndex;

            if (_prevSlideIndex - 3 >= 0)
            {
                newPrevSlideIndex = _prevSlideIndex - 3;
                PowerPointSlide prevSlide = this.GetCurrentPresentation().Slides[_prevSlideIndex - 3];
                AddSlideThumbnail(prevSlide, 0);
            }

            if (_prevSlideIndex - 2 >= 0)
            {
                PowerPointSlide prevSlide = this.GetCurrentPresentation().Slides[_prevSlideIndex - 2];
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
                PowerPointSlide prevSlide = this.GetCurrentPresentation().Slides[_prevSlideIndex - 1];
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
            int selectedIndex = SlideListBox.SelectedIndex;
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
            int newNextSlideIndex = _nextSlideIndex;

            if (_nextSlideIndex + 1 < this.GetCurrentPresentation().SlideCount)
            {
                PowerPointSlide nextSlide = this.GetCurrentPresentation().Slides[_nextSlideIndex + 1];
                newNextSlideIndex = _nextSlideIndex + 1;
                AddSlideThumbnail(nextSlide);
            }

            if (_nextSlideIndex + 2 < this.GetCurrentPresentation().SlideCount)
            {
                PowerPointSlide nextSlide = this.GetCurrentPresentation().Slides[_nextSlideIndex + 2];
                newNextSlideIndex = _nextSlideIndex + 2;
                AddSlideThumbnail(nextSlide);
            }

            if (_nextSlideIndex + 3 < this.GetCurrentPresentation().SlideCount)
            {
                PowerPointSlide nextSlide = this.GetCurrentPresentation().Slides[_nextSlideIndex + 3];
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

        private void SlideListBox_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListBoxItem item = ItemsControl.ContainerFromElement((ItemsControl)sender, (DependencyObject)e.OriginalSource)
                as ListBoxItem;
            if (item == null || item.Content == null)
            {
                return;
            }

            if (OnGotoSlide != null)
            {
                OnGotoSlide();
            }
        }
    }
}
