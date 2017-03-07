using System;
using System.Collections.Generic;

using Microsoft.Office.Core;

namespace PowerPointLabs.DataSources
{
    internal class DrawingLabData
    {
        public void AddPropertyChangedHandler(Action<string> callPropertyChanged)
        {
            _propertyChangedHandlers.Add(callPropertyChanged);
        }

        public void PropertyChanged(string propertyName)
        {
            foreach (var handler in _propertyChangedHandlers)
            {
                handler(propertyName);
            }
        }

        public bool IsHotkeysInitialised = false;

        public bool HotkeysEnabled = true;

        // Properties - Record / Apply Displacement
        public float ShiftValueX;
        public float ShiftValueY;
        public float ShiftValueRotation;
        public bool ShiftIncludePositionX = true;
        public bool ShiftIncludePositionY = true;
        public bool ShiftIncludeRotation = true;

        // Properties - Record / Apply Position
        public float SavedValueX;
        public float SavedValueY;
        public float SavedValueRotation;
        public bool SavedIncludePositionX = true;
        public bool SavedIncludePositionY = true;
        public bool SavedIncludeRotation = true;


        // Properties - Record / Apply Format
        public bool FormatSyncLineStyle = true;

        public bool FormatHasLine = true;
        public bool FormatIncludeHasLine = true;
        public int FormatLineColor = 0x000000;
        public bool FormatIncludeLineColor = true;
        public float FormatLineWeight = 5;
        public bool FormatIncludeLineWeight = true;
        public MsoLineDashStyle FormatLineDashStyle = MsoLineDashStyle.msoLineSolid;
        public bool FormatIncludeLineDashStyle = true;


        public bool FormatSyncFillStyle = true;

        public bool FormatHasFill = true;
        public bool FormatIncludeHasFill = true;
        public int FormatFillColor = 0xC07000;
        public bool FormatIncludeFillColor = true;


        public bool FormatSyncTextStyle = true;

        public string FormatText = "Text";
        public bool FormatIncludeText = false;
        public int FormatTextColor = 0x000000;
        public bool FormatIncludeTextColor = true;
        public float FormatTextFontSize = 18;
        public bool FormatIncludeTextFontSize = true;
        public string FormatTextFont = "Calibri (Body)";
        public bool FormatIncludeTextFont = true;
        public bool FormatTextWrap = false;
        public bool FormatIncludeTextWrap = true;
        public MsoAutoSize FormatTextAutoSize = MsoAutoSize.msoAutoSizeNone;
        public bool FormatIncludeTextAutoSize = true;


        public bool FormatSyncSize = false;

        public float FormatWidth = 0;
        public bool FormatIncludeWidth = true;
        public float FormatHeight = 0;
        public bool FormatIncludeHeight = true;

        // Properties - Anchor
        public DrawingsLabDataSource.Horizontal AnchorHorizontal = DrawingsLabDataSource.Horizontal.Center;
        public DrawingsLabDataSource.Vertical AnchorVertical = DrawingsLabDataSource.Vertical.Middle;

        private readonly List<Action<string>> _propertyChangedHandlers = new List<Action<string>>();
    }
}
