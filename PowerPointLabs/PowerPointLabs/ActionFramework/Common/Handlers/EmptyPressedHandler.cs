using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Common.Handlers
{
    class EmptyPressedHandler : PressedHandler
    {
        protected override bool GetPressed(string ribbonId)
        {
            return false;
        }
    }
}
