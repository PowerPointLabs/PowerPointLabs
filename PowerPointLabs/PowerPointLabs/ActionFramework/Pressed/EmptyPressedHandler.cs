using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Pressed
{
    class EmptyPressedHandler : PressedHandler
    {
        protected override bool GetPressed(string ribbonId, string ribbonTag)
        {
            return false;
        }
    }
}
