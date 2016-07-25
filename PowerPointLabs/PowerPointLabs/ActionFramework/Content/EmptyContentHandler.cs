using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Content
{
    class EmptyContentHandler : ContentHandler
    {
        protected override string GetContent(string ribbonId)
        {
            return "";
        }
    }
}
