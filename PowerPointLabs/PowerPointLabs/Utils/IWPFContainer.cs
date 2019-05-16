using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace PowerPointLabs.Utils
{
    /// <summary>
    /// Adapter interface for adaptive theme colors.
    /// </summary>
    public interface IWPFContainer
    {
        Control WpfControl { get; }
    }
}
