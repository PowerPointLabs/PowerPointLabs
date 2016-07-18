using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    class EmptyLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId, string ribbonTag)
        {
            return "";
        }
    }
}
