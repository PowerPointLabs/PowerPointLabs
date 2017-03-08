﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace PowerPointLabs.SyncLab.View
{
    class SyncFormatDialogItem : SyncFormatListItem
    {
        private SyncFormatDialogItem parent = null;
        private SyncFormatDialogItem[] children = null;
        private FormatTreeNode node = null;

        public SyncFormatDialogItem(FormatTreeNode node) : base()
        {
            checkBox.Checked += new RoutedEventHandler(CheckBoxCheckChange);
            checkBox.Unchecked += new RoutedEventHandler(CheckBoxCheckChange);
            this.node = node;
        }

        private void CheckBoxCheckChange(object sender, RoutedEventArgs e)
        {
            UpdateChecked();
        }

        public SyncFormatDialogItem ItemParent
        {
            get
            {
                return parent;
            }
            set
            {
                parent = value;
            }
        }

        public SyncFormatDialogItem[] ItemChildren
        {
            get
            {
                return children;
            }
            set
            {
                children = value;
            }
        }

        public new bool? IsChecked
        {
            get
            {
                return checkBox.IsChecked;
            }
            set
            {
                if (checkBox.IsChecked == value)
                {
                    return;
                }
                checkBox.IsChecked = value;
                node.IsChecked = value;
                UpdateChecked();
            }
        }

        private void UpdateChecked()
        {
            bool? value = checkBox.IsChecked;
            if (this.ItemChildren != null && value.HasValue)
            { // value = true/false
                foreach (SyncFormatDialogItem child in this.ItemChildren)
                {
                    child.IsChecked = value.Value;
                }
            }
            if (this.ItemParent != null)
            {
                this.ItemParent.UpdateChildrenChecked();
            }
        }

        private void UpdateChildrenChecked()
        {
            if (this.ItemChildren != null)
            {
                bool allFalse = true;
                bool allTrue = true;
                foreach (SyncFormatDialogItem child in this.ItemChildren)
                {
                    bool? childIsChecked = child.IsChecked;
                    if (!childIsChecked.HasValue || childIsChecked.Value)
                    {
                        allFalse = false;
                    }
                    if (!childIsChecked.HasValue || !childIsChecked.Value)
                    {
                        allTrue = false;
                    }
                }
                if (allFalse)
                {
                    IsChecked = false;
                }
                else if (allTrue)
                {
                    IsChecked = true;
                }
                else
                {
                    IsChecked = null;
                }
            }
        }
    }
}
