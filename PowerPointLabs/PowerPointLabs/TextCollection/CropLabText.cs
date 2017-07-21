namespace PowerPointLabs.TextCollection
{
    internal static class CropLabText
    {
        public const string ErrorUndefined = "Undefined error in '{0}'.";

        public const string ErrorSelectionIsInvalid = "You need to select at least {1} {2} before applying '{0}'.";
        public const string ErrorSelectionCountZero = "'{0}' requires at least one shape to be selected.";
        public const string ErrorSelectionNonPicture = "'{0}' only supports picture objects.";

        public const string ErrorSelectionMustBeShape = "'{0}' only supports shape objects.";
        public const string ErrorSelectionMustBePicture = "'{0}' only supports picture objects.";
        public const string ErrorSelectionMustBeShapeOrPicture = "'{0}' only supports shape or picture objects.";

        public const string ErrorNoShapeOverBoundary = "All selected objects are inside the slide boundary. No cropping was done.";
        public const string ErrorNoDimensionCropped = "All selected pictures are smaller than reference shape. No cropping was done.";
        public const string ErrorNoPaddingCropped = "All selected pictures have no transparent padding. No cropping was done.";
        public const string ErrorNoAspectRatioCropped = "All selected pictures are already in the given aspect ratio. No cropping was done.";

        public const string ErrorAspectRatioIsInvalid = "The given aspect ratio is invalid. Please enter positive numbers for the width to height ratio.";
    }
}
