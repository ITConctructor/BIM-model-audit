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
            //InitializeFileSystemObjects();
        }
        private void foldersItem_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {

        }

        //private void LoadCheckings(string Dep, Type[] checkings, DataGrid CheckingGrid)
        //{
        //    _checkings = new BindingList<CheckingTemplate>();
        //    for (int i = 0; i < checkings.Length; i++)
        //    {
        //        var checking = Activator.CreateInstance(Type.GetType(checkings[i].FullName)) as CheckingTemplate;
        //        //var tabItem = CheckingGrid.Parent as TabItem;
        //        if (checking is CheckingTemplate && checkings[i].Name != "CheckingTemplate")
        //        {
        //            if (Dep == checking.Dep)
        //            {
        //                _checkings.Add(checking as CheckingTemplate);
        //            }
        //        }
        //    }
        //    if (_checkings.Count > 0)
        //    {
        //        CheckingGrid.ItemsSource = _checkings;
        //    }
        //}

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            Type[] checkings = Assembly.GetExecutingAssembly().GetTypes().Where(t => t.Namespace == "Audit").ToArray();
            foreach (TabItem item in CheckingsTabControl.Items)
            {
                _checkings = new BindingList<CheckingTemplate>();
                DataGrid dataGrid = item.Content as DataGrid;
                //LoadCheckings(item.Header.ToString(), checkings, dataGrid);
                for (int i = 0; i < checkings.Length; i++)
                {
                    var checking = Activator.CreateInstance(Type.GetType(checkings[i].FullName)) as CheckingTemplate;
                    //var tabItem = CheckingGrid.Parent as TabItem;
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

        
        //[Serializable]
        //public abstract class PropertyNotifier : INotifyPropertyChanged
        //{
        //    public PropertyNotifier() : base() { }

        //    [field: NonSerialized]
        //    public event PropertyChangedEventHandler PropertyChanged;

        //    protected void OnPropertyChanged(string propertyName)
        //    {
        //        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        //    }
        //}
        //[Serializable]
        //public abstract class BaseObject : PropertyNotifier
        //{
        //    private IDictionary<string, object> m_values = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        //    public T GetValue<T>(string key)
        //    {
        //        var value = GetValue(key);
        //        return (value is T) ? (T)value : default(T);
        //    }

        //    private object GetValue(string key)
        //    {
        //        if (string.IsNullOrEmpty(key))
        //        {
        //            return null;
        //        }
        //        return m_values.ContainsKey(key) ? m_values[key] : null;
        //    }

        //    public void SetValue(string key, object value)
        //    {
        //        if (!m_values.ContainsKey(key))
        //        {
        //            m_values.Add(key, value);
        //        }
        //        else
        //        {
        //            m_values[key] = value;
        //        }
        //        OnPropertyChanged(key);
        //    }
        //}
        //[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        //public struct ShellFileInfo
        //{
        //    public IntPtr hIcon;

        //    public int iIcon;

        //    public uint dwAttributes;

        //    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        //    public string szDisplayName;

        //    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        //    public string szTypeName;
        //}
        //public static class Interop
        //{
        //    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        //    public static extern IntPtr SHGetFileInfo(string path,
        //        uint attributes,
        //        out ShellFileInfo fileInfo,
        //        uint size,
        //        uint flags);

        //    [DllImport("user32.dll", SetLastError = true)]
        //    [return: MarshalAs(UnmanagedType.Bool)]
        //    public static extern bool DestroyIcon(IntPtr pointer);
        //}
        //public enum FileAttribute : uint
        //{
        //    Directory = 16,
        //    File = 256
        //}
        //[Flags]
        //public enum ShellAttribute : uint
        //{
        //    LargeIcon = 0,              // 0x000000000
        //    SmallIcon = 1,              // 0x000000001
        //    OpenIcon = 2,               // 0x000000002
        //    ShellIconSize = 4,          // 0x000000004
        //    Pidl = 8,                   // 0x000000008
        //    UseFileAttributes = 16,     // 0x000000010
        //    AddOverlays = 32,           // 0x000000020
        //    OverlayIndex = 64,          // 0x000000040
        //    Others = 128,               // Not defined, really?
        //    Icon = 256,                 // 0x000000100  
        //    DisplayName = 512,          // 0x000000200
        //    TypeName = 1024,            // 0x000000400
        //    Attributes = 2048,          // 0x000000800
        //    IconLocation = 4096,        // 0x000001000
        //    ExeType = 8192,             // 0x000002000
        //    SystemIconIndex = 16384,    // 0x000004000
        //    LinkOverlay = 32768,        // 0x000008000 
        //    Selected = 65536,           // 0x000010000
        //    AttributeSpecified = 131072 // 0x000020000
        //}
        //public enum IconSize : short
        //{
        //    Small,
        //    Large
        //}
        //public enum ItemState : short
        //{
        //    Undefined,
        //    Open,
        //    Close
        //}
        //public enum ItemType
        //{
        //    Drive,
        //    Folder,
        //    File
        //}
        //public class FileSystemObjectInfo : BaseObject
        //{
        //    public FileSystemObjectInfo(DriveInfo drive)
        //        : this(drive.RootDirectory)
        //    {
        //    }

        //    public FileSystemObjectInfo(FileSystemInfo info)
        //    {
        //        if (this is DummyFileSystemObjectInfo)
        //        {
        //            return;
        //        }

        //        Children = new ObservableCollection<FileSystemObjectInfo>();
        //        FileSystemInfo = info;

        //        if (info is DirectoryInfo)
        //        {
        //            ImageSource = FolderManager.GetImageSource(info.FullName, ItemState.Close);
        //            AddDummy();
        //        }
        //        else if (info is FileInfo)
        //        {
        //            ImageSource = FileManager.GetImageSource(info.FullName);
        //        }

        //        PropertyChanged += new PropertyChangedEventHandler(FileSystemObjectInfo_PropertyChanged);
        //    }
        //    public event EventHandler BeforeExpand;

        //    public event EventHandler AfterExpand;

        //    public event EventHandler BeforeExplore;

        //    public event EventHandler AfterExplore;

        //    private void RaiseBeforeExpand()
        //    {
        //        BeforeExpand?.Invoke(this, EventArgs.Empty);
        //    }

        //    private void RaiseAfterExpand()
        //    {
        //        AfterExpand?.Invoke(this, EventArgs.Empty);
        //    }

        //    private void RaiseBeforeExplore()
        //    {
        //        BeforeExplore?.Invoke(this, EventArgs.Empty);
        //    }

        //    private void RaiseAfterExplore()
        //    {
        //        AfterExplore?.Invoke(this, EventArgs.Empty);
        //    }
        //    public ObservableCollection<FileSystemObjectInfo> Children
        //    {
        //        get { return base.GetValue<ObservableCollection<FileSystemObjectInfo>>("Children"); }
        //        private set { base.SetValue("Children", value); }
        //    }

        //    public ImageSource ImageSource
        //    {
        //        get { return base.GetValue<ImageSource>("ImageSource"); }
        //        private set { base.SetValue("ImageSource", value); }
        //    }

        //    public bool IsExpanded
        //    {
        //        get { return base.GetValue<bool>("IsExpanded"); }
        //        set { base.SetValue("IsExpanded", value); }
        //    }

        //    public FileSystemInfo FileSystemInfo
        //    {
        //        get { return base.GetValue<FileSystemInfo>("FileSystemInfo"); }
        //        private set { base.SetValue("FileSystemInfo", value); }
        //    }

        //    private DriveInfo Drive
        //    {
        //        get { return base.GetValue<DriveInfo>("Drive"); }
        //        set { base.SetValue("Drive", value); }
        //    }
        //    private void AddDummy()
        //    {
        //        this.Children.Add(new DummyFileSystemObjectInfo());
        //    }

        //    private bool HasDummy()
        //    {
        //        return !object.ReferenceEquals(this.GetDummy(), null);
        //    }

        //    private DummyFileSystemObjectInfo GetDummy()
        //    {
        //        var list = this.Children.OfType<DummyFileSystemObjectInfo>().ToList();
        //        if (list.Count > 0) return list.First();
        //        return null;
        //    }

        //    private void RemoveDummy()
        //    {
        //        this.Children.Remove(this.GetDummy());
        //    }
        //    private void ExploreDirectories()
        //    {
        //        if (Drive?.IsReady == false)
        //        {
        //            return;
        //        }
        //        if (FileSystemInfo is DirectoryInfo)
        //        {
        //            var directories = ((DirectoryInfo)FileSystemInfo).GetDirectories();
        //            foreach (var directory in directories.OrderBy(d => d.Name))
        //            {
        //                if ((directory.Attributes & FileAttributes.System) != FileAttributes.System &&
        //                    (directory.Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
        //                {
        //                    var fileSystemObject = new FileSystemObjectInfo(directory);
        //                    fileSystemObject.BeforeExplore += FileSystemObject_BeforeExplore;
        //                    fileSystemObject.AfterExplore += FileSystemObject_AfterExplore;
        //                    Children.Add(fileSystemObject);
        //                }
        //            }
        //        }
        //    }

        //    private void FileSystemObject_AfterExplore(object sender, EventArgs e)
        //    {
        //        RaiseAfterExplore();
        //    }

        //    private void FileSystemObject_BeforeExplore(object sender, EventArgs e)
        //    {
        //        RaiseBeforeExplore();
        //    }
        //    void FileSystemObjectInfo_PropertyChanged(object sender, PropertyChangedEventArgs e)
        //    {
        //        if (FileSystemInfo is DirectoryInfo)
        //        {
        //            if (string.Equals(e.PropertyName, "IsExpanded", StringComparison.CurrentCultureIgnoreCase))
        //            {
        //                RaiseBeforeExpand();
        //                if (IsExpanded)
        //                {
        //                    ImageSource = FolderManager.GetImageSource(FileSystemInfo.FullName, ItemState.Open);
        //                    if (HasDummy())
        //                    {
        //                        RaiseBeforeExplore();
        //                        RemoveDummy();
        //                        ExploreDirectories();
        //                        ExploreFiles();
        //                        RaiseAfterExplore();
        //                    }
        //                }
        //                else
        //                {
        //                    ImageSource = FolderManager.GetImageSource(FileSystemInfo.FullName, ItemState.Close);
        //                }
        //                RaiseAfterExpand();
        //            }
        //        }
        //    }
        //    private void ExploreFiles()
        //    {
        //        if (Drive?.IsReady == false)
        //        {
        //            return;
        //        }
        //        if (FileSystemInfo is DirectoryInfo)
        //        {
        //            var files = ((DirectoryInfo)FileSystemInfo).GetFiles();
        //            foreach (var file in files.OrderBy(d => d.Name))
        //            {
        //                if ((file.Attributes & FileAttributes.System) != FileAttributes.System &&
        //                    (file.Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
        //                {
        //                    Children.Add(new FileSystemObjectInfo(file));
        //                }
        //            }
        //        }
        //    }

        //}
        //public class ShellManager
        //{
        //    private static Icon Icon;

        //    public static Icon GetIcon(string path, ItemType type, IconSize iconSize, ItemState state)
        //    {
        //        var attributes = (uint)(type == ItemType.Folder ? FileAttribute.Directory : FileAttribute.File);
        //        var flags = (uint)(ShellAttribute.Icon | ShellAttribute.UseFileAttributes);

        //        if (type == ItemType.Folder && state == ItemState.Open)
        //        {
        //            flags = flags | (uint)ShellAttribute.OpenIcon;
        //        }
        //        if (iconSize == IconSize.Small)
        //        {
        //            flags = flags | (uint)ShellAttribute.SmallIcon;
        //        }
        //        else
        //        {
        //            flags = flags | (uint)ShellAttribute.LargeIcon;
        //        }

        //        var fileInfo = new ShellFileInfo();
        //        var size = (uint)Marshal.SizeOf(fileInfo);
        //        var result = Interop.SHGetFileInfo(path, attributes, out fileInfo, size, flags);

        //        if (result == IntPtr.Zero)
        //        {
        //            throw Marshal.GetExceptionForHR(Marshal.GetHRForLastWin32Error());
        //        }

        //        try
        //        {
        //            return (Icon)Icon.FromHandle(fileInfo.hIcon).Clone();
        //        }
        //        catch
        //        {
        //            throw;
        //        }
        //        finally
        //        {
        //            Interop.DestroyIcon(fileInfo.hIcon);
        //        }
        //    }
        //}
        //public static class FolderManager
        //{
        //    public static ImageSource GetImageSource(string directory, ItemState folderType)
        //    {
        //        return GetImageSource(directory, new System.Drawing.Size(16, 16), folderType);
        //    }

        //    public static ImageSource GetImageSource(string directory, System.Drawing.Size size, ItemState folderType)
        //    {
        //        using (var icon = ShellManager.GetIcon(directory, ItemType.Folder, IconSize.Large, folderType))
        //        {
        //            return Imaging.CreateBitmapSourceFromHIcon(icon.Handle,
        //                System.Windows.Int32Rect.Empty,
        //                BitmapSizeOptions.FromWidthAndHeight(size.Width, size.Height));
        //        }
        //    }
        //}
        //public static class FileManager
        //{
        //    public static ImageSource GetImageSource(string filename)
        //    {
        //        return GetImageSource(filename, new System.Drawing.Size(16, 16));
        //    }

        //    public static ImageSource GetImageSource(string filename, System.Drawing.Size size)
        //    {
        //        using (var icon = ShellManager.GetIcon(System.IO.Path.GetExtension(filename), ItemType.File, IconSize.Small, ItemState.Undefined))
        //        {
        //            return Imaging.CreateBitmapSourceFromHIcon(icon.Handle,
        //                System.Windows.Int32Rect.Empty,
        //                BitmapSizeOptions.FromWidthAndHeight(size.Width, size.Height));
        //        }
        //    }
        //}
        //private void InitializeFileSystemObjects()
        //{
        //    var drives = DriveInfo.GetDrives();
        //    DriveInfo.GetDrives().ToList().ForEach(drive =>
        //    {
        //        var fileSystemObject = new FileSystemObjectInfo(drive);
        //        fileSystemObject.BeforeExplore += FileSystemObject_BeforeExplore;
        //        fileSystemObject.AfterExplore += FileSystemObject_AfterExplore;
        //        treeView.Items.Add(fileSystemObject);
        //    });
        //}

        //private void FileSystemObject_AfterExplore(object sender, System.EventArgs e)
        //{
        //    Cursor = Cursors.Arrow;
        //}

        //private void FileSystemObject_BeforeExplore(object sender, System.EventArgs e)
        //{
        //    Cursor = Cursors.Wait;
        //}
        //internal class DummyFileSystemObjectInfo : FileSystemObjectInfo
        //{
        //    public DummyFileSystemObjectInfo()
        //        : base(new DirectoryInfo("DummyFileSystemObjectInfo"))
        //    {
        //    }
        //}
    }
}
