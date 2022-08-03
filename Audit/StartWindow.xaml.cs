﻿using System;
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
                System.Windows.Controls.DataGrid dataGrid = item.Content as System.Windows.Controls.DataGrid;
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

        private void AddFiles_Click(object sender, RoutedEventArgs e)
        {
            List<string> filesToAnalysAsText = new List<string>();
            foreach (TextBlock item in FilesToAnalys.Items)
            {
                filesToAnalysAsText.Add(item.Text);
            }
            TreeViewItem file = foldersItem.SelectedItem as TreeViewItem;
            TextBlock fileName = new TextBlock();
            fileName.Focusable = true;
            fileName.Text = file.Header as string;
            if (fileName.Text.EndsWith(".rvt") && !filesToAnalysAsText.Contains(fileName.Text))
            {
                FilesToAnalys.Items.Add(fileName);
                CommandLauncher.filesToAnalysPaths.Add(file.Tag as string);
                TextBlock activeFileName = new TextBlock();
                activeFileName.Text = file.Header as string;
                activeFileComboBox.Items.Add(activeFileName);
            }
        }

        private void RemoveFiles_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < FilesToAnalys.SelectedItems.Count; i = i)
            {
                FilesToAnalys.Items.Remove(FilesToAnalys.SelectedItems[i]);
            }
        }

        private void UpdateSelected_Click(object sender, RoutedEventArgs e)
        {
            TabItem selectedTabItem = CheckingsTabControl.SelectedItem as TabItem;
            DataGrid selectedDataGrid = selectedTabItem.Content as DataGrid;
            foreach (var item in selectedDataGrid.SelectedItems)
            {
                CheckingTemplate activeChecking = (CheckingTemplate)item;
                activeChecking.Run();
            }
        }

        private void UpdateAll_Click(object sender, RoutedEventArgs e)
        {
            foreach (TabItem item in CheckingsTabControl.Items)
            {
                DataGrid activeDataGrid = item.Content as DataGrid;
                foreach (var activeDataGridItem in activeDataGrid.Items)
                {
                    CheckingTemplate activeChecking = (CheckingTemplate)activeDataGridItem;
                    activeChecking.Run();
                }
            }
        }

        private void ReadLog()
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            StreamWriter writer = File.AppendText("");
        }

        private void WriteLog()
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            logFilePath.Text = folderBrowserDialog.SelectedPath;
        }
    }
}
