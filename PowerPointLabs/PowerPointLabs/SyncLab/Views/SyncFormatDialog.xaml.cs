using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.TextCollection;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.SyncLab.Views
{
    /// <summary>
    /// Interaction logic for SyncFormatDialog.xaml
    /// </summary>
    public partial class SyncFormatDialog : Window
    {
        public FormatTreeNode[] Formats { get; private set; }

        private string originalName;
        private Shape shape;

        public SyncFormatDialog(Shape shape) : this(shape, shape.Name, SyncFormatConstants.FormatCategories)
        {
        }

        public SyncFormatDialog(Shape shape, string formatName, FormatTreeNode[] formats)
        {
            InitializeComponent();
            this.shape = shape;

            formatName = formatName.Trim();
            if (SyncFormatUtil.IsValidFormatName(formatName))
            {
                originalName = formatName;
            }
            else
            {
                originalName = SyncLabText.DefaultFormatName;
            }
            originalName = formatName;
            ObjectName = originalName;
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
        
        private FormatTreeNode GetNodeWithFormatType(FormatTreeNode[] nodes, Type type)
        {
            List<FormatTreeNode> list = GetNodeWithFormatTypeHelper(nodes, type);
            
            if (list.Count == 0)
            {
                return null;
            }
            return list[0];
        }
        
        private List<FormatTreeNode> GetNodeWithFormatTypeHelper(FormatTreeNode[] nodes, Type type)
        {
            List<FormatTreeNode> list = new List<FormatTreeNode>();
            foreach (FormatTreeNode node in nodes)
            {
                if (node.IsFormatNode)
                {
                    if (node.Format.GetType() == type)
                    {
                        list.Add(node);
                    }
                }
                else
                {
                    list.AddRange(GetNodeWithFormatTypeHelper(node.ChildrenNodes, type));
                }
            }
            return list;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;

            ShowWarningMessageForMixedStylePerspective();
        }

        /**
         * Check if custom perspective shadow was used & show a warning if so
         * We cannot handle it accurately, see ShadowEffectFormat.cs for more information
         */
        private void ShowWarningMessageForMixedStylePerspective()
        {
            // check if custom perspective shadow was used,
            FormatTreeNode shadowNode = GetNodeWithFormatType(Formats, typeof(ShadowEffectFormat));
            if (shadowNode == null)
            {
                return;
            }

            bool shadowNodeIsChecked = shadowNode.IsChecked != null 
                                       && shadowNode.IsChecked.Value;

            if (ShadowEffectFormat.MightHaveCustomPerspectiveShadow(shape) 
                && shadowNodeIsChecked)
            {
                MessageBox.Show(SyncLabText.WarningSyncPerspectiveShadow, SyncLabText.WarningDialogTitle);
            }
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
                    return this.originalName;
                }
            }
            set
            {
                nameTextBox.Text = value;
            }
        }
    }
}
