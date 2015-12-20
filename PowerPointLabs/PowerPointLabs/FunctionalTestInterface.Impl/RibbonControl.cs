
using Microsoft.Office.Core;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    class RibbonControl : IRibbonControl
    {
        public RibbonControl(string id)
        {
            this.Id = id;
        }

        public string Id { get; set; }

        public object Context { get; set; }

        public string Tag { get; set; }
    }
}
