﻿using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Models;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.WPF.Observable;
using System.Windows.Media;

namespace PowerPointLabs.PictureSlidesLab.View
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

        private bool _isLoadNextSlide = true;

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
            var currentSlide = this.GetCurrentSlide();
            if (currentSlide.Index - 2 >= 0)
            {
                _prevSlideIndex = currentSlide.Index - 2;
                var prevSlide = this.GetCurrentPresentation().Slides[currentSlide.Index - 2];
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
                var nextSlide = this.GetCurrentPresentation().Slides[currentSlide.Index];
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

        /// <summary>
        /// This method can only be called after dialog is fully initialized
        /// </summary>
        /// <param name="imageItems"></param>
        /// <param name="title"></param>
        public SlideSelectionDialog Init(System.Collections.Generic.List<ImageItem> imageItems, string title)
        {
            DialogTitleProperty.Text = title;
            _isLoadNextSlide = false;
            Dispatcher.Invoke(new Action(() =>
            {
                SlideList.Clear();
            }));

            foreach (var imageItem in imageItems)
            {
                SlideList.Add(imageItem);
            }

            SelectCurrentSlide();
            SlideListBox.ScrollIntoView(SlideListBox.SelectedItem);

            if (SlideListBox.SelectedIndex == 0)
            {
                _prevSlideIndex = SlideListBox.SelectedIndex;
                _nextSlideIndex = SlideListBox.SelectedIndex + 1;
            }
            else if (SlideListBox.SelectedIndex == SlideList.Count - 1)
            {
                _prevSlideIndex = SlideListBox.SelectedIndex - 1;
                _nextSlideIndex = SlideListBox.SelectedIndex;
            }
            else
            {
                _prevSlideIndex = SlideListBox.SelectedIndex - 1;
                _nextSlideIndex = SlideListBox.SelectedIndex + 1;
            }

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

        private void SelectCurrentSlide()
        {
            foreach (var slide in SlideList)
            {
                if (slide.Tooltip.Contains("Current"))
                {
                    SlideListBox.SelectedItem = slide;
                    return;
                }
            }

            SlideListBox.SelectedItem = SlideList[0];
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
            return (int)(this.GetCurrentPresentation().SlideWidth / this.GetCurrentPresentation().SlideHeight * PreviewHeight);
        }

        private Visual GetDescendantByType(Visual element, Type type)
        {
            if (element == null)
            {
                return null;
            }

            if (element.GetType() == type)
            {
                return element;
            }

            Visual foundElement = null;

            if (element is FrameworkElement)
            {
                (element as FrameworkElement).ApplyTemplate();
            }

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(element); i++)
            {
                var visual = VisualTreeHelper.GetChild(element, i) as Visual;
                foundElement = GetDescendantByType(visual, type);

                if (foundElement != null)
                {
                    break;
                }
            }

            return foundElement;
        }

        private void GotoSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (OnGotoSlide != null)
            {
                OnGotoSlide();
            }
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
                var selectedItem = (ImageItem) SlideListBox.SelectedItem;
                var selectedString = selectedItem.Tooltip;
                if (selectedString.StartsWith("(Current) "))
                {
                    selectedString = selectedString.Remove(0, 10);
                }
                if (selectedString.StartsWith("Slide "))
                {
                    selectedString = selectedString.Remove(0, 6);
                    SelectedSlide = Int32.Parse(selectedString);
                }
                else
                {
                    SelectedSlide = SlideListBox.SelectedIndex;
                }
            }
        }

        private void SlideListBox_OnLoaded(object sender, RoutedEventArgs e)
        {
            if (SlideListBox.SelectedItem != null)
            {
                var scrollViewer = GetDescendantByType(SlideListBox, typeof(ScrollViewer)) as ScrollViewer;
                var itemCenter = scrollViewer.ExtentWidth / SlideListBox.Items.Count * (SlideListBox.SelectedIndex + 0.5);
                scrollViewer.ScrollToHorizontalOffset(Math.Min(scrollViewer.ExtentWidth - scrollViewer.ViewportWidth,
                    Math.Max(0, itemCenter - scrollViewer.ViewportWidth / 2)));
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
                var prevSlide = this.GetCurrentPresentation().Slides[_prevSlideIndex - 3];
                AddSlideThumbnail(prevSlide, 0);
            }

            if (_prevSlideIndex - 2 >= 0)
            {
                var prevSlide = this.GetCurrentPresentation().Slides[_prevSlideIndex - 2];
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
                var prevSlide = this.GetCurrentPresentation().Slides[_prevSlideIndex - 1];
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
                if (_isLoadNextSlide)
                {
                    LoadNextSlides();
                }
            }
        }

        private void LoadNextSlides(bool isToSelectNextSlide = true)
        {
            var newNextSlideIndex = _nextSlideIndex;

            if (_nextSlideIndex + 1 < this.GetCurrentPresentation().SlideCount)
            {
                var nextSlide = this.GetCurrentPresentation().Slides[_nextSlideIndex + 1];
                newNextSlideIndex = _nextSlideIndex + 1;
                AddSlideThumbnail(nextSlide);
            }

            if (_nextSlideIndex + 2 < this.GetCurrentPresentation().SlideCount)
            {
                var nextSlide = this.GetCurrentPresentation().Slides[_nextSlideIndex + 2];
                newNextSlideIndex = _nextSlideIndex + 2;
                AddSlideThumbnail(nextSlide);
            }

            if (_nextSlideIndex + 3 < this.GetCurrentPresentation().SlideCount)
            {
                var nextSlide = this.GetCurrentPresentation().Slides[_nextSlideIndex + 3];
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
            var item = ItemsControl.ContainerFromElement((ItemsControl)sender, (DependencyObject)e.OriginalSource)
                as ListBoxItem;
            if (item == null || item.Content == null) return;

            if (OnGotoSlide != null)
            {
                OnGotoSlide();
            }
        }
    }
}
