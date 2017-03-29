using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SyncLab.View
{
    class SyncFormatPaneItem : SyncFormatListItem
    {
        ItemsControl parent;
        RadioButton radioButton;
        Shape shape = null;
        FormatTreeNode[] formats = null;

        public SyncFormatPaneItem(ItemsControl parent, Shape shape, FormatTreeNode[] formats) : base()
        {
            this.parent = parent;
            this.shape = shape;
            this.formats = formats;
            radioButton = new RadioButton();
            radioButton.Margin = checkBox.Margin;
            radioButton.HorizontalAlignment = checkBox.HorizontalAlignment;
            radioButton.VerticalAlignment = checkBox.VerticalAlignment;
            radioButton.Width = checkBox.Width;
            radioButton.Height = checkBox.Height;
            grid.Children.Add(radioButton);
            grid.Children.Remove(checkBox);
            radioButton.Checked += new RoutedEventHandler(RadioBoxChecked);
            this.MouseDoubleClick += OnMouseDoubleClick;
        }

        public Shape FormatShape
        {
            get
            {
                return shape;
            }
        }

        public FormatTreeNode[] Formats
        {
            get
            {
                return formats;
            }
            set
            {
                formats = value;
            }
        }

        public new bool? IsChecked
        {
            get
            {
                return radioButton.IsChecked;
            }
            set
            {
                radioButton.IsChecked = value;
            }
        }

        private void RadioBoxChecked(object sender, RoutedEventArgs e)
        {
            foreach (Object obj in parent.Items)
            {
                SyncFormatPaneItem item = (SyncFormatPaneItem)obj;
                if (item != this)
                {
                    item.IsChecked = false;
                }
            }
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            SyncFormatDialog dialog = new SyncFormatDialog(shape, this.Text, this.formats);
            dialog.ObjectName = this.Text;
            bool? result = dialog.ShowDialog();
            if (!result.HasValue || !(bool)result)
            {
                return;
            }
            this.formats = dialog.Formats;
            this.Text = dialog.ObjectName;
        }
    }
}
