using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml;
using System.Xml.Serialization;
using DataGrid = System.Windows.Controls.DataGrid;
using Excel = Microsoft.Office.Interop.Excel;

namespace Audit
{
    public class ApplicationViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public class RelayCommand : ICommand
        {
            private Action _execute;

            public event EventHandler CanExecuteChanged
            {
                add { CommandManager.RequerySuggested += value; }
                remove { CommandManager.RequerySuggested -= value; }
            }

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public void Execute(object parameter)
            {
                _execute();
            }

            public RelayCommand(Action execute)
            {
                _execute = execute;
            }
        }

        public ApplicationViewModel(StartWindow win)
        {
            _win = win;
            LoadLastLaunchData();
            FirstLaunchDataInitializing();
            ReadLog();
        }

        #region Данные для передачи в StartWindow

        private StartWindow _win;

        public ObservableCollection<RvtFileInfo> PreanalysFiles { get; set; } = new ObservableCollection<RvtFileInfo>();

        private string _logFilesPath = Properties.Settings.Default.folderToSaveLog;
        public string LogFilesPath { get => _logFilesPath; set { SetProperty(ref _logFilesPath, value); } }

        private RvtFileInfo _selectedFile;
        public RvtFileInfo SelectedFile 
        { 
            get => _selectedFile; 
            set 
            {
                RvtFileInfo buf = value as RvtFileInfo;
                SetProperty(ref _selectedFile, buf);
                _win.CommonCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ОБЩ").ToList();
                _win.ARCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "АР").ToList();
                _win.KRCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "КР").ToList();
                _win.EOMCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ЭОМ").ToList();
                _win.VKCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ВК").ToList();
                _win.OVCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ОВ").ToList();
                _win.SSCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "СС").ToList();
            } 
        }

        private CheckingTemplate _selectedChecking;
        public CheckingTemplate SelectedChecking
        {
            get { return _selectedChecking; }
            set 
            { 
                CheckingTemplate buf = value as CheckingTemplate;
                SetProperty(ref _selectedChecking, buf);
                _win.ResultsGrid.ItemsSource = _selectedChecking?.ElementCheckingResults;
            }
        }

        private ReportSettings _reportsSettings = new ReportSettings() { ReportType = "Все тесты", ReportFormat = "Excel" };
        public ReportSettings ReportsSettings
        { 
            get => _reportsSettings; 
            set { SetProperty(ref _reportsSettings, value); } 
        }
        #endregion

        #region Методы
        /// <summary>
        /// Загружает файловую структуру с ПК для просмотрщика TreeView
        /// </summary>
        private static List<object> GetFileStructure(Bitmap DriveIcon, Bitmap FolderIcon, string type = "")
        {
            List<object> FileStructure = new List<object>();
            foreach (string drive in Directory.GetLogicalDrives())
            {
                List<object> Elements = new List<object>();
                BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                            DriveIcon.GetHbitmap(),
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                ImageSource DriveImage = mediumImage as ImageSource;
                BitmapSource mediumFolderImage = Imaging.CreateBitmapSourceFromHBitmap(
                            FolderIcon.GetHbitmap(),
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                ImageSource FolderImage = mediumFolderImage as ImageSource;
                try
                {
                    foreach (string Folder in Directory.GetDirectories(drive))
                    {
                        Elements.Add((FolderImage, Folder, ExploreFolders(Folder, FolderIcon, "rvt")));
                    }
                    foreach (string File in Directory.GetFiles(drive))
                    {
                        if (File.EndsWith(type))
                        {
                            System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(File);
                            ImageSourceConverter imageConverter = new ImageSourceConverter();
                            ImageSource FileImage = Imaging.CreateBitmapSourceFromHIcon(icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                            Elements.Add((FileImage, File));
                        }
                    }
                }
                catch (Exception)
                {
                    Elements.Add((DriveImage, drive));
                }
                FileStructure.Add((DriveImage, drive, Elements));
            }
            return FileStructure;
        }

        /// <summary>
        /// Возвращает все папки и файлы, содержащиеся внутри папки
        /// </summary>
        private static List<object> ExploreFolders(string Folder, Bitmap FolderIcon, string type = "")
        {
            List<object> Elements = new List<object>();
            BitmapSource mediumImage = Imaging.CreateBitmapSourceFromHBitmap(
                        FolderIcon.GetHbitmap(),
                        IntPtr.Zero,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());
            ImageSource FolderImage = mediumImage as ImageSource;
            foreach (string folder in Directory.GetDirectories(Folder))
            {
                try
                {
                    if (Directory.GetDirectories(folder).Length > 0 || Directory.GetFiles(Folder).Length > 0)
                    {
                        List<object> subFolders = ExploreFolders((string)folder, FolderIcon);
                        Elements.Add((FolderImage, folder, subFolders));
                    }
                    else
                    {
                        Elements.Add((FolderImage, folder));
                    }
                }
                catch (Exception)
                {
                    Elements.Add((FolderImage, folder));
                }
            }
            foreach (string File in Directory.GetFiles(Folder))
            {
                if (File.EndsWith(type))
                {
                    System.Drawing.Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(File);
                    ImageSourceConverter imageConverter = new ImageSourceConverter();
                    ImageSource FileImage = Imaging.CreateBitmapSourceFromHIcon(icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    Elements.Add((FileImage, File));
                }
            }
            return Elements;
        }

        /// <summary>
        /// Добавляет выбранные в файловой структуре файлы в список для последующего анализа
        /// </summary>
        private void AddFilesToPreanalysList()
        {
            List<string> filesToAnalysAsText = new List<string>();
            foreach (RvtFileInfo item in _win.FilesToAnalys.Items)
            {
                filesToAnalysAsText.Add(item.Name);
            }
            TreeViewItem file = _win.foldersItem.SelectedItem as TreeViewItem;
            string Name = file.Header as string;
            string Path = file.Tag as string;
            RvtFileInfo selectedFile = new RvtFileInfo() { Name = Name , Path = Path };
            if (Name.EndsWith(".rvt") && !filesToAnalysAsText.Contains(Name))
            {
                PreanalysFiles.Add(selectedFile);
            }
        }

        /// <summary>
        /// Удаляет файл из списка файлов для анализа
        /// </summary>
        private void RemoveFilesFromPreanalysList()
        {
            for (int i = 0; i < _win.FilesToAnalys.SelectedItems.Count; i = i)
            {
                RvtFileInfo FileToRemove = _win.FilesToAnalys.SelectedItems[i] as RvtFileInfo;
                PreanalysFiles.Remove(FileToRemove);
            }
        }

        /// <summary>
        /// Изменение значение свойства. Отличается от стандартного функционала тем, что при изменении значения свойства в коде нет необходимости менять его в OnPropertyChanged
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="storage">Значение свойства до изменения</param>
        /// <param name="value">Передаваемое значение</param>
        /// <param name="propertyName">Имя свойства</param>
        /// <returns></returns>
        private protected bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(storage, value)) return false;
            storage = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        /// <summary>
        /// Выбирает в файловой структуре ПК папку для сохранения результатов проверок, впоследствии использующихся приложением
        /// </summary>
        private void SelectFolderToSaveLog()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            LogFilesPath = folderBrowserDialog.SelectedPath;
            Properties.Settings.Default.folderToSaveLog = LogFilesPath;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Загружает из настроек приложения данные прошлого запуска: анализированные файлы и результаты проверок
        /// </summary>
        private void LoadLastLaunchData()
        {
            string loadedLastAnalysFiles = Properties.Settings.Default.lastCheckedFiles;
            string[] loadedLastAnalysFilesArray = loadedLastAnalysFiles.Split('|');
            string loadedLastAnalysFilesPaths = Properties.Settings.Default.lastCheckedFilesPaths;
            string[] loadedLastAnalysFilesPathsArray = loadedLastAnalysFilesPaths.Split('|');
            PreanalysFiles.Clear();
            for (int i = 0; i < loadedLastAnalysFilesArray.Count(); i++)
            {
                if (loadedLastAnalysFilesArray[i] != "")
                {
                    RvtFileInfo File = new RvtFileInfo() { Name = loadedLastAnalysFilesArray[i], Path = loadedLastAnalysFilesPathsArray[i] };
                    PreanalysFiles.Add(File);
                }
            }
            string reportSettings = Properties.Settings.Default.ReportSettings;
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(ReportSettings));
            using(StringReader reader = new StringReader(reportSettings))
            {
                if (reportSettings != "")
                {
                    ReportsSettings = (ReportSettings)xmlSerializer.Deserialize(reader);
                }
            }

        }

        /// <summary>
        /// Сохраняет результаты проверок для последующей работы приложения
        /// </summary>
        private void WriteLog()
        {
            string logFilesPaths = "";
            foreach (RvtFileInfo File in PreanalysFiles)
            {
                string logFileName = File.Name;
                string directoryPath = Properties.Settings.Default.folderToSaveLog + "\\log_" + logFileName;
                directoryPath = directoryPath.Replace(".rvt", "");
                string fullFilePath = directoryPath + "\\log_" + logFileName.Substring(0, logFileName.LastIndexOf(".")) + "_" + DateTime.Now.ToString().Replace(' ', '_').Replace(':', '.') + ".xml";
                logFilesPaths = logFilesPaths + fullFilePath + "|";
                Directory.CreateDirectory(directoryPath);
                Serialize(File.CheckingResults, fullFilePath);
                string text = System.IO.File.ReadAllText(fullFilePath);
                text = text.Substring(0, text.LastIndexOf(">") + 1);
                System.IO.File.WriteAllText(fullFilePath, text, Encoding.UTF8);
            }
            
            Properties.Settings.Default.logFilesPaths = logFilesPaths;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Создает xml файл по указанному пути
        /// </summary>
        public static void Serialize(List<FileCheckingResult> result, string file)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<FileCheckingResult>));
            string xml;
            System.Xml.XmlWriterSettings writerSets = new System.Xml.XmlWriterSettings();
            writerSets.NewLineOnAttributes = true;
            writerSets.Indent = true;
            writerSets.ConformanceLevel = ConformanceLevel.Document;
            using (XmlWriter writer = XmlWriter.Create(file, writerSets))
            {
                xmlSerializer.Serialize(writer, result);
                xml = writer.ToString();
            }
            File.AppendAllText(file, xml, Encoding.UTF8);
        }

        /// <summary>
        /// Выполняет все проверки, сохраняет их результаты во внешний файл и сохраняет выбор файлов для последующего запуска
        /// </summary>
        private void UpdateAll()
        {
            for (int i = 0; i < PreanalysFiles.Count; i++)
            {
                _win.activeFileComboBox.SelectedIndex = i;
                foreach (CheckingTemplate Checking in PreanalysFiles[i].CheckingResults[0].Checkings)
                {
                    UpdateChecking(Checking, PreanalysFiles[i].Path);
                }
            }
            UpdateSettings();
            WriteLog();
        }

        /// <summary>
        /// Выполняет выбранные проверки, сохраняет их результаты во внешний файл и сохраняет выбор файлов для последующего запуска
        /// </summary>
        private void UpdateSelected()
        {
            foreach (RvtFileInfo file in PreanalysFiles)
            {
                TabItem selectedTabItem = _win.CheckingsTabControl.SelectedItem as TabItem;
                DataGrid selectedDataGrid = selectedTabItem.Content as DataGrid;
                foreach (var item in selectedDataGrid.SelectedItems)
                {
                    CheckingTemplate activeChecking = (CheckingTemplate)item;
                    foreach (CheckingTemplate checking in file.CheckingResults[0].Checkings)
                    {
                        if (checking.Name == activeChecking.Name)
                        {
                            UpdateChecking(checking, file.Path);
                        }
                    }
                }
            }
            UpdateSettings();
            WriteLog();
        }

        /// <summary>
        /// Общий для UpdateSelected и UpdateAll метод запуска проверки и обновления результатов
        /// </summary>
        private void UpdateChecking(CheckingTemplate Checking, string filePath)
        {
            Checking.LastRun = System.DateTime.Now.ToString();
            Checking.Running.Run(filePath, Checking.ElementCheckingResults);
            Checking.Amount = Checking.ElementCheckingResults.Count.ToString();
            Checking.Created = Checking.ElementCheckingResults.Where(t => t.Status == "Созданная").ToList().Count.ToString();
            Checking.Active = Checking.ElementCheckingResults.Where(t => t.Status == "Активная").ToList().Count.ToString();
            foreach (ElementCheckingResult result in Checking.ElementCheckingResults)
            {
                if (result.Status == null)
                {
                    result.Status = "Исправленная";
                }
            }
            Checking.Corrected = Checking.ElementCheckingResults.Where(t => t.Status == "Исправленная").ToList().Count.ToString();
            Checking.Checked = Checking.ElementCheckingResults.Where(t => t.Status == "Проверенная").ToList().Count.ToString();
        }

        /// <summary>
        /// Обновляет данные последнего запуска для последующего показа в приложении
        /// </summary>
        private void UpdateSettings()
        {
            string lastAnalysFilesString = "";
            string lastAnalysFilesPathsString = "";
            foreach (RvtFileInfo fileToSaveInList in PreanalysFiles)
            {
                lastAnalysFilesString = lastAnalysFilesString + fileToSaveInList.Name + "|";
                lastAnalysFilesPathsString = lastAnalysFilesPathsString + fileToSaveInList.Path + "|";
            }
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(ReportSettings));
            string reportSettings = "";
            using(StringWriter writer = new StringWriter())
            {
                xmlSerializer.Serialize(writer, ReportsSettings);
                reportSettings = writer.ToString();
            }
            Properties.Settings.Default.lastCheckedFiles = lastAnalysFilesString;
            Properties.Settings.Default.lastCheckedFilesPaths = lastAnalysFilesPathsString;
            Properties.Settings.Default.ReportSettings = reportSettings;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Заполняет массив данных пустыми значениями результатов проверок, если плагин запускается впервые и в лог ничего не записано
        /// </summary>
        private void FirstLaunchDataInitializing()
        {
            Type[] checkings = Assembly.GetExecutingAssembly().GetTypes().Where(t => t.Namespace == "Audit.Checkings").ToArray();
            foreach (RvtFileInfo FileInfo in PreanalysFiles)
            {
                if (FileInfo.CheckingResults.Count == 0)
                {
                    FileInfo.CheckingResults.Add(new FileCheckingResult());
                    foreach (TabItem item in _win.CheckingsTabControl.Items)
                    {
                        DataGrid dataGrid = item.Content as DataGrid;
                        foreach (Type checking in checkings)
                        {
                            CheckingTemplate Checking = Activator.CreateInstance(Type.GetType(checking.FullName)) as CheckingTemplate;
                            if (item.Header.ToString() == Checking.Dep)
                            {
                                CheckingTemplate NewChecking = new CheckingTemplate();
                                NewChecking.Dep = Checking.Dep;
                                NewChecking.Name = Checking.Name;
                                NewChecking.Running = Checking;
                                FileInfo.CheckingResults[0].Checkings.Add(NewChecking);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Считывает результаты предыдущих проверок
        /// </summary>
        private void ReadLog()
        {
            string[] logFilesPaths = Properties.Settings.Default.logFilesPaths.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string item in logFilesPaths)
            {
                foreach (RvtFileInfo fileInfo in PreanalysFiles)
                {
                    if (item.IndexOf(fileInfo.Name.Replace(".rvt", "")) != -1)
                    {
                        fileInfo.CheckingResults[0] = Deserialize(item);
                    }
                }
            }
        }
        
        /// <summary>
        /// Считывает xml файл
        /// </summary>
        public static FileCheckingResult Deserialize(string file)
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<FileCheckingResult>));
            List<FileCheckingResult> results = new List<FileCheckingResult>();
            FileCheckingResult result = new FileCheckingResult();
            if (File.Exists(file))
            {
                try
                {
                    using (XmlReader reader = XmlReader.Create(file))
                    {
                        results = (List<FileCheckingResult>)xmlSerializer.Deserialize(reader);
                        result = results[results.Count - 1];
                    }
                }
                catch (Exception)
                {

                }
            }
            return result;
        }

        /// <summary>
        /// Общий для всех классов проверок метод добавления результата проверки элемента в список
        /// </summary>
        public static void AddElementCheckingResult(ElementCheckingResult newResult, BindingList<ElementCheckingResult> results)
        {
            if (results.Remove(newResult))
            {
                newResult.Status = "Активная";
                results.Add(newResult);
            }
            else
            {
                newResult.Status = "Созданная";
                results.Add(newResult);
            }
        }

        /// <summary>
        /// Создает результаты проверок в формат excel
        /// </summary>
        private void CreateExcel()
        {
            TaskDialog dialog = new TaskDialog("Test");
            dialog.MainContent = "Excel created";
            dialog.Show();
        }
        #endregion

        #region Комманды
        private RelayCommand _addFilesToPreanalysListCommand;
        public RelayCommand AddFilesToPreanalysListCommand
        {
            get { return _addFilesToPreanalysListCommand ?? (_addFilesToPreanalysListCommand = new RelayCommand(AddFilesToPreanalysList)); }
        }

        private RelayCommand _removeFilesFromPreanalysListCommand;
        public RelayCommand RemoveFilesFromPreanalysListCommand
        {
            get { return _removeFilesFromPreanalysListCommand ?? (_removeFilesFromPreanalysListCommand = new RelayCommand(RemoveFilesFromPreanalysList)); }
        }

        private RelayCommand _selectFolderToSaveLogCommand;
        public RelayCommand SelectFolderToSaveLogCommand
        {
            get { return _selectFolderToSaveLogCommand ?? (_selectFolderToSaveLogCommand = new RelayCommand(SelectFolderToSaveLog)); }
        }

        private RelayCommand _updateAllCommand;
        public RelayCommand UpdateAllCommand
        {
            get { return _updateAllCommand ?? (_updateAllCommand = new RelayCommand(UpdateAll)); }
        }

        private RelayCommand _updateSelectedCommand;
        public RelayCommand UpdateSelectedCommand
        {
            get { return _updateSelectedCommand ?? (_updateSelectedCommand = new RelayCommand(UpdateSelected)); }
        }

        private RelayCommand _createExcelCommand;
        public RelayCommand CreateExcelCommand
        {
            get { return _createExcelCommand ?? (_createExcelCommand = new RelayCommand(CreateExcel)); }
        }
        #endregion

        #region Вспомогательные классы для хранения данных
        public class RvtFileInfo
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public List<FileCheckingResult> CheckingResults { get; set; } = new List<FileCheckingResult>();
            public override string ToString()
            {
                return Name;
            }
        }

        public class ReportSettings
        {
            public class ReportFields
            {
                public bool Name { get; set; }
                public bool Status { get; set; }
                public bool Id { get; set; }
                public bool Comment { get; set; }
                public bool Time { get; set; }
            }
            public class ReportStatuses
            {
                public bool Created { get; set; }
                public bool Active { get; set; }
                public bool Checked { get; set; }
                public bool Corrected { get; set; }
            }
            public ReportFields Fields { get; set; } = new ReportFields();
            public ReportStatuses Statuses { get; set; } = new ReportStatuses();
            public string ReportType { get; set; }
            public string ReportFormat { get; set; }
        }
        #endregion
    }
}
