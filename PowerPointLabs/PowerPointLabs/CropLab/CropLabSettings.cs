﻿namespace PowerPointLabs.CropLab
{
    public enum AnchorPosition
    {
        TopLeft,
        Top,
        TopRight,
        MiddleLeft,
        Middle,
        MiddleRight,
        BottomLeft,
        Bottom,
        BottomRight,
        Reference
    }

    internal static class CropLabSettings
    {
        public static AnchorPosition AnchorPosition = AnchorPosition.Reference;

        public static float GetAnchorX()
        {
            switch (AnchorPosition)
            {
                case AnchorPosition.TopLeft:
                case AnchorPosition.MiddleLeft:
                case AnchorPosition.BottomLeft:
                    return 0;
                case AnchorPosition.Top:
                case AnchorPosition.Middle:
                case AnchorPosition.Bottom:
                case AnchorPosition.Reference:
                    return 0.5F;
                case AnchorPosition.TopRight:
                case AnchorPosition.MiddleRight:
                case AnchorPosition.BottomRight:
                    return 1;
            }
            return 0;
        }

        public static float GetAnchorY()
        {
            switch (AnchorPosition)
            {
                case AnchorPosition.TopLeft:
                case AnchorPosition.Top:
                case AnchorPosition.TopRight:
                    return 0;
                case AnchorPosition.MiddleLeft:
                case AnchorPosition.Middle:
                case AnchorPosition.MiddleRight:
                case AnchorPosition.Reference:
                    return 0.5F;
                case AnchorPosition.BottomLeft:
                case AnchorPosition.Bottom:
                case AnchorPosition.BottomRight:
                    return 1;
            }
            return 0;
        }
    }
    
}
