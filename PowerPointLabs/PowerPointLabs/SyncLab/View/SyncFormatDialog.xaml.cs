﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace PowerPointLabs.SyncLab.View
{
    /// <summary>
    /// Interaction logic for SyncFormatDialog.xaml
    /// </summary>
    public partial class SyncFormatDialog : Window
    {

        FormatTreeNode[] formats = null;

        public SyncFormatDialog(Shape shape) : this(shape, SyncFormatConstants.FormatCategories)
        {
        }

        public SyncFormatDialog(Shape shape, FormatTreeNode[] formats)
        {
            InitializeComponent();
            this.formats = formats;
            foreach (FormatTreeNode format in formats)
            {
                Object treeItem = DialogItemFromFormatTree(shape, format);
                if (treeItem != null)
                {
                    treeView.Items.Add(treeItem);
                }
            }   
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
                        else
                        {
                            throw new Exception("");
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
            this.DialogResult = true;
        }

        public FormatTreeNode[] Formats
        {
            get
            {
                return formats;
            }
        }
    }
}
