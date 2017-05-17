using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace PowerPointLabs.SyncLab.View
{
    /// <summary>
    /// Interaction logic for SyncFormatDialog.xaml
    /// </summary>
    public partial class SyncFormatDialog : Window
    {
        public delegate void OkButtonEventHandler(SyncFormatDialog dialog);
        public event OkButtonEventHandler OkButtonClick;
        public string OriginalName { get; private set; }
        public FormatTreeNode[] Formats { get; private set; }
        public Shape Shape { get; private set; }

        private SyncPaneWPF parent;

        public SyncFormatDialog(SyncPaneWPF parent, Shape shape, string formatName, FormatTreeNode[] formats)
        {
            InitializeComponent();
            this.parent = parent;
            this.Shape = shape;

            formatName = formatName.Trim();
            if (SyncFormatUtil.IsValidFormatName(formatName))
            {
                OriginalName = formatName;
            }
            else
            {
                OriginalName = TextCollection.SyncLabDefaultFormatName;
            }
            OriginalName = formatName;
            ObjectName = OriginalName;
            this.Formats = formats;
            foreach (FormatTreeNode format in formats)
            {
                Object treeItem = DialogItemFromFormatTree(shape, format);
                if (treeItem != null)
                {
                    treeView.Items.Add(treeItem);
                }
            }
            ScrollToTop();
        }

        private Object DialogItemFromFormatTree(Shape shape, FormatTreeNode node)
        {
            if (node.Format != null)
            {
                SyncFormatDialogItem result = new SyncFormatDialogItem(node);
                if (!node.Format.CanCopy(shape))
                {
                    return null;
                }
                result.Text = node.Name;
                result.IsChecked = node.IsChecked;
                result.Image = node.Format.DisplayImage(shape);
                return result;
            }
            else
            {
                SyncFormatDialogItem header = new SyncFormatDialogItem(node);
                header.Text = node.Name;
                header.IsChecked = node.IsChecked;
                TreeViewItem result = new TreeViewItem();
                result.Header = header;
                List<SyncFormatDialogItem> children = new List<SyncFormatDialogItem>();
                foreach (FormatTreeNode childNode in node.ChildrenNodes)
                {
                    Object treeItem = DialogItemFromFormatTree(shape, childNode);
                    if (treeItem != null)
                    {
                        if (treeItem is SyncFormatDialogItem)
                        {
                            children.Add((SyncFormatDialogItem)treeItem);
                        }
                        else if (treeItem is TreeViewItem)
                        {
                            children.Add((SyncFormatDialogItem)((TreeViewItem)treeItem).Header);
                        }
                        result.Items.Add(treeItem);
                    }
                }
                header.ItemChildren = children.ToArray<SyncFormatDialogItem>();
                foreach (SyncFormatDialogItem child in children)
                {
                    child.ItemParent = header;
                }
                result.ExpandSubtree();
                if (result.Items.Count == 0)
                {
                    return null;
                }
                return result;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            OkButtonClick(this);
            parent.Dialog = null;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            parent.Dialog = null;
            this.Close();
        }

        private void ScrollToTop()
        {
            if (!treeView.Items.IsEmpty)
            {
                (treeView.Items[0] as TreeViewItem).BringIntoView();
            }
        }

        public string ObjectName
        {
            get
            {
                if (SyncFormatUtil.IsValidFormatName(nameTextBox.Text))
                {
                    return nameTextBox.Text.Trim();
                }
                else
                {
                    return this.OriginalName;
                }
            }
            set
            {
                nameTextBox.Text = value;
            }
        }
    }
}
