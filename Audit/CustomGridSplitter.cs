using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Audit
{
    internal class CustomGridSplitter : GridSplitter
    {
        public enum SplitterDirectionEnum
        {
            Horizontal,
            Vertical
        }
        public SplitterDirectionEnum SplitterDirection { get; set; }
        public int MinimumDistanceFromEdge { get; set; }
        public System.Windows.Point originPoint { get; set; }
    }
}
