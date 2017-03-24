namespace PowerPointLabs.CropLab
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
        BottomRight
    }

    internal static class CropLabSettings
    {
        private static AnchorPosition anchorPosition = AnchorPosition.Middle;

        public static AnchorPosition AnchorPosition
        {
            get { return anchorPosition; }
            set { anchorPosition = value; }
        }
    }
}
