using Audit.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Audit.View
{
    /// <summary>
    /// Interaction logic for ParameterCheckingWindow.xaml
    /// </summary>
    public partial class ParameterCheckingWindow : Window
    {
        public ParameterCheckingWindow(ParameterCheckingViewModel viewModel)
        {
            InitializeComponent();
            ViewModel = viewModel;
            DataContext = ViewModel;
        }
        public ParameterCheckingViewModel ViewModel { get; private set; }
    }
}
