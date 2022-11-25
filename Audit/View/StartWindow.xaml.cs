using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
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
using Image = System.Drawing.Image;
using Microsoft.Win32;
using System.Windows.Forms;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using DataGrid = System.Windows.Controls.DataGrid;
using System.Xml.Serialization;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Xml.Linq;

namespace Audit
{
    public partial class StartWindow : Window
    {
        public StartWindow()
        {
            InitializeComponent();
            ViewModel = new ApplicationViewModel(this);
            DataContext = ViewModel;
        }
        private ApplicationViewModel ViewModel { get; set; }

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

        private void folder_Expanded(object sender, RoutedEventArgs e)
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
                    foreach (string s in Directory.GetFiles(item.Tag.ToString()))
                    {
                        if (s.EndsWith(".rvt"))
                        {
                            TreeViewItem file = new TreeViewItem();
                            file.Header = s.Substring(s.LastIndexOf("\\") + 1);
                            file.Tag = s;
                            file.FontWeight = FontWeights.Normal;
                            item.Items.Add(file);
                        }
                    }
                }
                catch (Exception) { }
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            GetDrivers();
            activeFileComboBox.SelectedIndex = 0;
            activeGrid.SelectedIndex = 0;

        }

        private void CheckingsTabControl_LostFocus(object sender, RoutedEventArgs e)
        {
            if (CheckingsTabControl.SelectedItem == null)
            {
                CheckingsTabControl.SelectedIndex = 0;
                TabItem activeTab = CheckingsTabControl.SelectedItem as TabItem;
                System.Windows.Controls.DataGrid activeGrid = activeTab.Content as System.Windows.Controls.DataGrid;
                activeGrid.SelectedIndex = 0;
            }
        }

        private void splitterCenter_MouseDown(object sender, MouseButtonEventArgs e)
        {
            splitterCenter.originPoint = e.GetPosition(Window.GetWindow(this));
        }

        private void splitterCenter_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
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

        private System.Windows.Controls.DataGrid activeGrid = new System.Windows.Controls.DataGrid();

        private void CheckingsTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabItem activeTab = CheckingsTabControl.SelectedItem as TabItem;
            activeGrid = activeTab.Content as System.Windows.Controls.DataGrid;
        }

        private void ResultStatusSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ViewModel.SelectedChecking.UpdateCounts();
        }

        private void BaseWindow_Closed(object sender, EventArgs e)
        {
            ViewModel.WriteLog();
        }
    }
}
