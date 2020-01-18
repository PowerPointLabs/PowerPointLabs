using Microsoft.Office.Core;
using PowerPointLabs.SyncLab;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    public static class ShapeTypeExtension
    {
        public static bool IsNormalShape(this Shape shape)
        {
            return shape.Type == MsoShapeType.msoAutoShape ||
                shape.Type == MsoShapeType.msoLine ||
                shape.Type == MsoShapeType.msoPicture ||
                shape.Type == MsoShapeType.msoTextBox;
        }

        public static bool IsGroupShape(this Shape shape)
        {
            return shape.Type == MsoShapeType.msoGroup;
        }

        public static bool IsPlaceholderSyncable(this Shape shape)
        {
            if (shape.Type != MsoShapeType.msoPlaceholder)
            {
                return false;
            }
            Microsoft.Office.Interop.PowerPoint.Shapes templateShapes =
                SyncFormatUtil.GetTemplateShapes();
            return ShapeUtil.CanCopyMsoPlaceHolder(shape, templateShapes);
        }

    }
}
