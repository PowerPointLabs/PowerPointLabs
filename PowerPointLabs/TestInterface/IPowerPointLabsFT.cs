using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestInterface
{
    public interface IPowerPointLabsFT
    {
        IPowerPointLabsFeatures GetFeatures();
        IPowerPointOperations GetOperations();
    }
}
