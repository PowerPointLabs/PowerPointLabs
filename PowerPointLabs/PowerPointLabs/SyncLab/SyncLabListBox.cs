using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.SyncLab
{
    public partial class SyncLabListBox : ListView
    {

        public static readonly int MAX_LIST_SIZE = 50;
        List<ListViewItem> formatList = new List<ListViewItem>();

        public SyncLabListBox()
        {
            InitializeComponent();
            this.CheckBoxes = true;
            this.View = View.LargeIcon;
            this.LargeImageList = new ImageList();
            this.LargeImageList.ImageSize = ObjectFormat.DISPLAY_IMAGE_SIZE;
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
        }

        public void AddFormat(ObjectFormat format)
        {
            // Add thumbnail to list
            string imageKey = GetNextImageKey();
            this.LargeImageList.Images.Add(imageKey, format.DisplayImage);
            // Add item to list
            ListViewItem newItem = new ListViewItem(format.DisplayText, imageKey);
            newItem.Tag = format;
            formatList.Insert(0, newItem);
            // Remove excess items
            if (formatList.Count > MAX_LIST_SIZE)
            {
                int countToRemove = formatList.Count - MAX_LIST_SIZE;
                formatList.RemoveRange(MAX_LIST_SIZE, countToRemove);
            }
            // Update displayed items
            this.BeginUpdate();
            this.Items.Clear();
            this.Items.AddRange(formatList.ToArray());
            this.EndUpdate();
        }

        public ObjectFormat GetFormat(int index)
        {
            return (ObjectFormat)this.Items[index].Tag;
        }

        public void RemoveFormat(int index)
        {
            string imageKey = this.Items[index].ImageKey;
            this.Items.RemoveAt(index);
            this.LargeImageList.Images.RemoveByKey(imageKey);
        }

        private int curImageIndex = 0;
        private string GetNextImageKey()
        {
            string key;
            do
            { // Find new index for the image
                key = (curImageIndex++).ToString();
            }
            while (this.LargeImageList.Images.ContainsKey(key));
            return key;
        }
    }
}
