using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Pressed
{
    class EmptyPressedHandler : PressedHandler
    {
        protected override bool GetPressed(string ribbonId)
        {
            return false;
        }
    }
}
