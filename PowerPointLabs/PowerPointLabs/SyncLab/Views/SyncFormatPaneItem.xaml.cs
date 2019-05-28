using System;
using System.Drawing;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.Views
{
    /// <summary>
    /// Interaction logic for SyncFormatPaneItem.xaml
    /// </summary>
    public partial class SyncFormatPaneItem : UserControl
    {

        private Bitmap image;

        private SyncPaneWPF parent;
        private string shapeKey = null;
        private SyncLabShapeStorage shapeStorage;
        private FormatTreeNode[] formats = null;

        #region Constructors

        public SyncFormatPaneItem(SyncPaneWPF parent, string shapeKey,
                    SyncLabShapeStorage shapeStorage, FormatTreeNode[] formats)
        {
            InitializeComponent();
            this.parent = parent;
            this.shapeKey = shapeKey;
            this.shapeStorage = shapeStorage;
            this.formats = formats;
            editImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.SyncLabEditButton);
            pasteImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.SyncLabPasteButton);
            deleteImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.SyncLabDeleteButton);

            UpdateToolTipBody();
        }

        #endregion

        #region Getters and Setters

        public string FormatShapeKey
        {
            get
            {
                return shapeKey;
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

        public Bitmap Image
        {
            get
            {
                return image;
            }
            set
            {
                image = value;
                UpdateImage();
            }
        }

        public bool FormatShapeExists
        {
            get
            {
                return shapeStorage.GetShape(shapeKey) != null;
            }
        }

        private void UpdateImage()
        {
            // if image isn't set, fill the control with the label
            if (image == null)
            {
                imageBox.Visibility = Visibility.Hidden;
                col1.Width = new GridLength(0);
                return;
            }
            else
            {
                BitmapSource source = CommonUtil.CreateBitmapSource(image);
                imageBox.Source = source;
                imageBox.Visibility = Visibility.Visible;
                col1.Width = new GridLength(60);
            }
        }

        public string Text
        {
            get
            {
                return textBlock.Text;
            }
            set
            {
                textBlock.Text = value;
                toolTipName.Text = value;
            }
        }

        #endregion

        #region Helper Functions

        private void UpdateToolTipBody()
        {
            StringBuilder toolTipBodyText = new StringBuilder();
            foreach (FormatTreeNode format in formats)
            {
                toolTipBodyText.Append(GetNamesOfCheckedNodes(format.Name, format));
            }
            toolTipBody.Text = toolTipBodyText.ToString().Trim("\n".ToCharArray());
        }

        private string GetNamesOfCheckedNodes(string rootName, FormatTreeNode node)
        {
            if (node.IsChecked ?? false)
            {
                return (node.Name.Equals(rootName) ? "" : rootName +
                    SyncFormatConstants.FormatNameSeparator) + node.Name + "\n";
            }
            StringBuilder result = new StringBuilder();
            foreach (FormatTreeNode child in node.ChildrenNodes ?? new FormatTreeNode[] { })
            {
                result.Append(GetNamesOfCheckedNodes(rootName, child));
            }
            return result.ToString();
        }

        private void ApplyFormatToSelected()
        {
            Shape formatShape = shapeStorage.GetShape(shapeKey);
            if (formatShape == null)
            {
                MessageBox.Show(SyncLabText.ErrorShapeDeleted, SyncLabText.ErrorDialogTitle);
                parent.ClearInvalidFormats();
            }
            this.StartNewUndoEntry();
            parent.ApplyFormats(formats, formatShape);
        }

        #endregion

        #region Event Handlers

        private void PasteButton_Click(object sender, RoutedEventArgs e)
        {
            ApplyFormatToSelected();
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            Shape shape = shapeStorage.GetShape(shapeKey);
            parent.Dialog = new SyncFormatDialog(shape, Text, formats);
            parent.Dialog.ObjectName = this.Text;
            bool? result = parent.Dialog.ShowThematicDialog();
            if (!result.HasValue || !(bool)result)
            {
                return;
            }
            this.formats = parent.Dialog.Formats;
            this.Text = parent.Dialog.ObjectName;
            parent.Dialog = null;
            UpdateToolTipBody();
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            parent.RemoveFormatItem(this);
        }

        private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ApplyFormatToSelected();
        }

        #endregion

    }
}
