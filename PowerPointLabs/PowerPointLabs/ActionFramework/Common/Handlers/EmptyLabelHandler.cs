using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Handlers
{
    class EmptyLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return "";
        }
    }
}
