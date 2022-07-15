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
    public partial class StartWindow : Window
    {
        private BindingList<CheckingTemplate> _checkings;
        public StartWindow()
        {
            InitializeComponent();
        }

        private void Load_Checkings()
        {
            Type[] checkings = Assembly.GetExecutingAssembly().GetTypes().Where(t => t.Namespace == "Audit").ToArray();
            foreach (TabItem item in CheckingsTabControl.Items)
            {
                _checkings = new BindingList<CheckingTemplate>();
                DataGrid dataGrid = item.Content as DataGrid;
                for (int i = 0; i < checkings.Length; i++)
                {
                    var checking = Activator.CreateInstance(Type.GetType(checkings[i].FullName)) as CheckingTemplate;
                    if (checking is CheckingTemplate && checkings[i].Name != "CheckingTemplate")
                    {
                        if (item.Header.ToString() == checking.Dep)
                        {
                            _checkings.Add(checking as CheckingTemplate);
                        }
                    }
                }
                if (_checkings.Count > 0)
                {
                    dataGrid.ItemsSource = _checkings;
                }
            }
        }

        private void GetDrivers()
        {
            foreach (string s in Directory.GetLogicalDrives())
            {
                TreeViewItem item = new TreeViewItem();
                TreeViewItem dummyNode = new TreeViewItem();
                dummyNode.Header = "dummyNode";
                item.Header = s;
                item.Tag = s;
                item.FontWeight = FontWeights.Normal;
                item.Items.Add(dummyNode);
                item.Expanded += new RoutedEventHandler(folder_Expanded);
                foldersItem.Items.Add(item);
            }
        }

        void folder_Expanded(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = (TreeViewItem)sender;
            TreeViewItem smallItem = item.Items[0] as TreeViewItem;
            
            if (item.Items.Count == 1 && smallItem.Header.ToString() == "dummyNode")
            {
                item.Items.Clear();
                try
                {
                    foreach (string s in Directory.GetDirectories(item.Tag.ToString()))
                    {
                        TreeViewItem dummyNode = new TreeViewItem();
                        dummyNode.Header = "dummyNode";
                        TreeViewItem subitem = new TreeViewItem();
                        subitem.Header = s.Substring(s.LastIndexOf("\\") + 1);
                        subitem.Tag = s;
                        subitem.FontWeight = FontWeights.Normal;
                        subitem.Items.Add(dummyNode);
                        subitem.Expanded += new RoutedEventHandler(folder_Expanded);
                        item.Items.Add(subitem);
                    }
                }
                catch (Exception) { }
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GetDrivers();
            Load_Checkings();
        }

        private void CheckingsTabControl_LostFocus(object sender, RoutedEventArgs e)
        {
            if (CheckingsTabControl.SelectedItem == null)
            {
                CheckingsTabControl.SelectedIndex = 0;
                TabItem tabItem = CheckingsTabControl.SelectedItem as TabItem;
                DataGrid dataGrid = tabItem.Content as DataGrid;
                dataGrid.SelectedIndex = 0;
            }
        }

        private void splitterCenter_MouseDown(object sender, MouseButtonEventArgs e)
        {
            splitterCenter.originPoint = e.GetPosition(Window.GetWindow(this));
        }

        private void splitterCenter_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                Window pwindow = Window.GetWindow(this);
                System.Windows.Point newPoint = e.GetPosition(pwindow);
                if (splitterCenter.SplitterDirection == CustomGridSplitter.SplitterDirectionEnum.Horizontal)
                {
                    if (newPoint.Y > splitterCenter.originPoint.Y)
                    {
                        if (newPoint.Y >= pwindow.ActualHeight - splitterCenter.MinimumDistanceFromEdge)
                        {
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        if (newPoint.Y > pwindow.ActualHeight - (splitterCenter.MinimumDistanceFromEdge + splitterCenter.Width))
                        {
                            e.Handled = true;
                        }
                    }
                }
                else
                {
                    if (newPoint.X > splitterCenter.originPoint.X)
                    {
                        if (newPoint.X >= pwindow.ActualWidth - splitterCenter.MinimumDistanceFromEdge)
                        {
                            e.Handled = true;
                        }
                    }
                    else
                    {
                        if (newPoint.X > pwindow.ActualWidth - (splitterCenter.MinimumDistanceFromEdge + splitterCenter.Width))
                        {
                            e.Handled = true;
                        }
                    }
                }
            }
        }
    }
}
