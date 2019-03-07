﻿using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.TooltipsLab.Views
{
    /// <summary>
    /// Interaction logic for TooltipsLabSettingsShapeEntry.xaml
    /// </summary>
    public partial class TooltipsLabSettingsShapeEntry : UserControl
    {
        private MsoAutoShapeType type;

        #region Constructors

        public TooltipsLabSettingsShapeEntry(MsoAutoShapeType type, Bitmap image)
        {
            InitializeComponent();
            Type = type;
            imageBox.Source = Imaging.CreateBitmapSourceFromHBitmap(
                image.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }

        #endregion

        #region Properties

        public MsoAutoShapeType Type
        {
            get
            {
                return type;
            }
            set
            {
                type = value;
                string nameForDisplay = value.ToString().Replace(
                    TooltipsLabConstants.ShapeNameHeader, "");
                textBlock.Text = nameForDisplay;
                ToolTip = nameForDisplay;
            }
        }

        #endregion
    }
}