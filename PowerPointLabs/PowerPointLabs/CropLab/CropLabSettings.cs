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
        public static AnchorPosition AnchorPosition = AnchorPosition.Middle;
    }
}
