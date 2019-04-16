using System;

using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointLabsFT : MarshalByRefObject, IPowerPointLabsFT
    {
        public static bool IsFunctionalTestOn;
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
