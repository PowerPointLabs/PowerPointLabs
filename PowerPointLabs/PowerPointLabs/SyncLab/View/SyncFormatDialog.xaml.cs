using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PowerPointLabs.SyncLab.View
{
    /// <summary>
    /// Interaction logic for SyncFormatDialog.xaml
    /// </summary>
    public partial class SyncFormatDialog : Window
    {

        private bool checkboxChanging = false;
        public SyncFormatDialog()
        {
            InitializeComponent();
            KeyValuePair<string, string[]>[] items = new KeyValuePair<string, string[]>[]
            {
                new KeyValuePair<string, string[]>(
                        "Text",
                        new string[]
                        {
                                "Font",
                                "Font Size",
                                "Font Color",
                                "Bold",
                                "Italics",
                                "Underline",
                                "Shadow",
                                "Strikethrow",
                                "Character spacing",
                                "Line spacing",
                                "Alignment"
                        }
                    ),
                new KeyValuePair<string, string[]>(
                        "Fill",
                        new string[]
                        {
                                "Fill"
                        }
                    ),
                new KeyValuePair<string, string[]>(
                        "Line",
                        new string[]
                        {
                                "Start arrow",
                                "End arrow",
                                "Weight",
                                "Style",
                                "Fill"
                        }
                    ),
                new KeyValuePair<string, string[]>(
                        "Effect",
                        new string[]
                        {
                                "Shadow",
                                "Reflection",
                                "Glow",
                                "Soft Edge",
                                "Bevel",
                                "3D Rotation"
                        }
                    ),
                new KeyValuePair<string, string[]>(
                        "Position/Size",
                        new string[]
                        {
                                "X",
                                "Y",
                                "Width",
                                "Height"
                        }
                    ),
            };
            System.Drawing.Bitmap b = new System.Drawing.Bitmap(50, 50);
            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(b);
            g.FillRectangle(System.Drawing.Brushes.DarkBlue, 0, 0, b.Width, b.Height);
            foreach (KeyValuePair<string, string[]> item in items)
            {
                TreeViewItem header = new TreeViewItem();
                SyncFormatListItem i = new SyncFormatListItem();
                List<SyncFormatListItem> children = new List<SyncFormatListItem>();
                i.Text = item.Key;
                i.Image = b;
                RoutedEventHandler headerCheckChange = new System.Windows.RoutedEventHandler(
                    (object sender, RoutedEventArgs e) =>
                    {
                        if (checkboxChanging)
                        {
                            return;
                        }
                        checkboxChanging = true;
                        bool? checkStatus = ((SyncFormatListItem)((Grid)((CheckBox)sender).Parent).Parent).IsChecked;
                        foreach (SyncFormatListItem obj in children)
                        {
                            obj.IsChecked = checkStatus;
                        }
                        checkboxChanging = false;
                    }
                );
                i.checkBox.Checked += headerCheckChange;
                i.checkBox.Unchecked += headerCheckChange;
                i.checkBox.Indeterminate += headerCheckChange;
                header.Header = i;
                foreach (string format in item.Value)
                {
                    SyncFormatListItem j = new SyncFormatListItem();
                    j.Text = format;
                    j.Image = b;
                    RoutedEventHandler checkChange = new System.Windows.RoutedEventHandler(
                        (object sender, RoutedEventArgs e) =>
                        {
                            if (checkboxChanging)
                            {
                                return;
                            }
                            checkboxChanging = true;
                            bool allChecked = true;
                            bool allUnchecked = true;
                            foreach (SyncFormatListItem subItem in children)
                            {
                                if (subItem.IsChecked.Value)
                                {
                                    allUnchecked = false;
                                }
                                else
                                {
                                    allChecked = false;
                                }
                            }
                            if (allChecked)
                            {
                                i.IsChecked = true;
                            }
                            else if (allUnchecked)
                            {
                                i.IsChecked = false;
                            }
                            else
                            {
                                i.IsChecked = null;
                            }
                            checkboxChanging = false;
                        }
                    );
                    j.checkBox.Checked += checkChange;
                    j.checkBox.Unchecked += checkChange;
                    j.checkBox.Indeterminate += checkChange;
                    header.Items.Add(j);
                    children.Add(j);
                }
                header.ExpandSubtree();
                treeView.Items.Add(header);
            }
        }

        private void CheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (checkboxChanging)
            {
                return;
            }
            checkboxChanging = true;
            CheckBox checkBox = (CheckBox)sender;
            SyncFormatListItem item = (SyncFormatListItem)((Grid)checkBox.Parent).Parent;
            if (((TreeViewItem)item.Parent).Parent is TreeViewItem)
            {
                // Recalculate parent
                bool allChecked = true;
                bool allUnchecked = true;
                foreach (object obj in ((TreeViewItem)item.Parent).Items)
                {
                    SyncFormatListItem subItem = (SyncFormatListItem)obj;
                    if (subItem.IsChecked.Value)
                    {
                        allUnchecked = false;
                    }
                    else
                    {
                        allChecked = false;
                    }
                }
                SyncFormatListItem parentItem = (SyncFormatListItem)((TreeViewItem)item.Parent).Header;
                if (allChecked)
                {
                    parentItem.IsChecked = true;
                }
                else if (allUnchecked)
                {
                    parentItem.IsChecked = false;
                }
                else
                {
                    parentItem.IsChecked = null;
                }
            }
            else
            {
                TreeViewItem parent = (TreeViewItem)item.Parent;
                foreach (Object obj in parent.Items)
                {
                    ((SyncFormatListItem)obj).IsChecked = item.IsChecked;
                }
            }
            checkboxChanging = false;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}
