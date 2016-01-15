using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    class PowerPointLabsFT : MarshalByRefObject, IPowerPointLabsFT
    {
        private static IPowerPointLabsFeatures features = new PowerPointLabsFeatures();
        private static IPowerPointOperations op = new PowerPointOperations();

        public IPowerPointLabsFeatures GetFeatures()
        {
            return features;
        }

        public IPowerPointOperations GetOperations()
        {
            return op;
        }
    }
}
