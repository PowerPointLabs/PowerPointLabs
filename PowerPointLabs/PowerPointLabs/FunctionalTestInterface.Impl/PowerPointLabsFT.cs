using System;

using PowerPointLabs.FunctionalTestInterface.Windows;

using TestInterface;
using TestInterface.Windows;

namespace PowerPointLabs.FunctionalTestInterface.Impl
{
    [Serializable]
    class PowerPointLabsFT : MarshalByRefObject, IPowerPointLabsFT
    {
        public static bool IsFunctionalTestOn;
        private static IPowerPointLabsFeatures features = new PowerPointLabsFeatures();
        private static IPowerPointOperations op = new PowerPointOperations();
        private static IWindowStackManager windowStackManager = new WindowStackManager();

        public IPowerPointLabsFeatures GetFeatures()
        {
            return features;
        }

        public IPowerPointOperations GetOperations()
        {
            return op;
        }

        public IWindowStackManager GetWindowStackManager()
        {
            return windowStackManager;
        }
    }
}
