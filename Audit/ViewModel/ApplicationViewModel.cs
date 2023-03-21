using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Xml;
using System.Xml.Serialization;
using DataGrid = System.Windows.Controls.DataGrid;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using Application = System.Windows.Forms.Application;
using Audit.View;
using Audit.Model;
using Audit.ViewModel;
using static Audit.Model.Utils;

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

        private DataBase Model { get; set; }

        public ApplicationViewModel(DataBase model)
        {
            View = new StartWindow(this);
            Model = model;
            LoadLastLaunchSettings();
            ReadLog();
            View.Show();
        }

        #region Данные для передачи в StartWindow

        private StartWindow View;

        public ObservableCollection<RvtFileInfo> PreanalysFiles
        {
            get
            {
                return Model.PreanalysFiles;
            }
            set
            {
                Model.PreanalysFiles = value;
            }
        }

        private string _logFilesPath = CommandLauncher.resultsPath;
        public string LogFilesPath { get => _logFilesPath; set { SetProperty(ref _logFilesPath, value); } }

        private bool _updateForSelectedFile;
        public bool UpdateForSelectedFile
        {
            get { return _updateForSelectedFile; }
            set
            {
                SetProperty(ref _updateForSelectedFile, value);
            }
        }

        private RvtFileInfo _selectedFile;
        public RvtFileInfo SelectedFile
        {
            get => _selectedFile;
            set
            {
                RvtFileInfo buf = value as RvtFileInfo;
                SetProperty(ref _selectedFile, buf);
                View.CommonCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ОБЩ").ToList();
                View.ARCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "АР").ToList();
                View.KRCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "КР").ToList();
                View.EOMCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ЭОМ").ToList();
                View.VKCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ВК").ToList();
                View.OVCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "ОВ").ToList();
                View.SSCheckingGrid.ItemsSource = _selectedFile?.CheckingResults[0].Checkings.Where(t => t.Dep == "СС").ToList();
                SelectedChecking = _selectedFile?.CheckingResults[0].Checkings[0];
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
                View.ResultsGrid.ItemsSource = _selectedChecking?.ElementCheckingResults;
                View.resultStatusesColumn.ItemsSource = ResultStatuses;
                if (buf != null && buf.ResultType == CheckingResultType.ElementsList)
                {
                    View.ResultsGrid.Visibility = System.Windows.Visibility.Visible;
                    View.MessageTextBlock.Visibility = System.Windows.Visibility.Hidden;
                }
                else if (buf != null && buf.ResultType == CheckingResultType.Message)
                {
                    View.ResultsGrid.Visibility = System.Windows.Visibility.Hidden;
                    View.MessageTextBlock.Visibility = System.Windows.Visibility.Visible;
                }
            }
        }

        private ReportSettings _reportsSettings = new ReportSettings() { Type = "Все тесты", Format = "Excel" };
        public ReportSettings ReportsSettings
        {
            get => _reportsSettings;
            set { SetProperty(ref _reportsSettings, value); }
        }

        public List<string> ReportTypes { get; set; } = new List<string>() { "Все тесты", "Выбранные тесты" };

        public List<string> ReportFormats { get; set; } = new List<string>() { "Excel", "HTML" };

        public static List<string> ResultStatuses { get; set; } = new List<string>() { "Созданная", "Активная", "Проверенная", "Исправленная" };

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
            List<string> preanalysFilesAsText = new List<string>();
            foreach (RvtFileInfo item in PreanalysFiles)
            {
                preanalysFilesAsText.Add(item.Name);
            }
            TreeViewItem file = View.foldersItem.SelectedItem as TreeViewItem;
            string Name = file.Header as string;
            string Path = file.Tag as string;
            RvtFileInfo selectedFile = new RvtFileInfo(Name, Path);
            if (Name.EndsWith(".rvt") && !preanalysFilesAsText.Contains(Name))
            {
                Model.PreanalysFiles.Add(selectedFile);
            }
        }

        /// <summary>
        /// Удаляет файл из списка файлов для анализа
        /// </summary>
        private void RemoveFilesFromPreanalysList()
        {
            for (int i = 0; i < View.FilesToAnalys.SelectedItems.Count; i = i)
            {
                RvtFileInfo FileToRemove = View.FilesToAnalys.SelectedItems[i] as RvtFileInfo;
                Model.PreanalysFiles.Remove(FileToRemove);
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
        public bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(storage, value)) return false;
            storage = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        /// <summary>
        /// Загружает из настроек приложения данные прошлого запуска: анализированные файлы и результаты проверок
        /// </summary>
        private void LoadLastLaunchSettings()
        {
            string[] loadedLastAnalysFilesArray = Properties.Settings.Default.lastCheckedFiles.Split('|');
            string[] loadedLastAnalysFilesPathsArray = Properties.Settings.Default.lastCheckedFilesPaths.Split('|');
            Model.PreanalysFiles.Clear();
            for (int i = 0; i < loadedLastAnalysFilesArray.Count(); i++)
            {
                if (loadedLastAnalysFilesArray[i] != "")
                {
                    RvtFileInfo File = new RvtFileInfo(loadedLastAnalysFilesArray[i], loadedLastAnalysFilesPathsArray[i]);
                    Model.PreanalysFiles.Add(File);
                }
            }
            string serializedReportSettings = Properties.Settings.Default.ReportSettings;
            if (serializedReportSettings != "")
            {
                ReportsSettings = Deserialize<ReportSettings>(serializedReportSettings);
            }

        }

        /// <summary>
        /// Считывает результаты предыдущих проверок
        /// </summary>
        private void ReadLog()
        {
            string[] logFilesPaths = Properties.Settings.Default.logFilesPaths.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string path in logFilesPaths)
            {
                foreach (RvtFileInfo fileInfo in Model.PreanalysFiles)
                {
                    foreach (CheckingTemplate checking in fileInfo.CheckingResults[0].Checkings)
                    {
                        checking.Status = CheckingStatus.NotLaunched;
                    }
                    if (path.Contains(fileInfo.Name.Replace(".rvt", "")) && File.Exists(path))
                    {
                        string serializedFileInfo = System.IO.File.ReadAllText(path);
                        fileInfo.CheckingResults = Deserialize<List<FileCheckingResult>>(serializedFileInfo);
                        foreach (CheckingTemplate checking in fileInfo.CheckingResults[fileInfo.CheckingResults.Count-1].Checkings)
                        {
                            checking.Status = CheckingStatus.NotUpdated;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Сохраняет результаты проверок для последующей работы приложения
        /// </summary>
        public void WriteLog()
        {
            string logFilesPaths = "";
            foreach (RvtFileInfo file in Model.PreanalysFiles)
            {
                string logFileName = file.Name;
                string directoryPath = CommandLauncher.resultsPath + "\\log_" + logFileName;
                directoryPath = directoryPath.Replace(".rvt", "");
                string fullFilePath = directoryPath + "\\log_" + logFileName.Substring(0, logFileName.LastIndexOf(".")) + "_" + DateTime.Now.ToString().Replace(' ', '_').Replace(':', '.') + ".xml";
                logFilesPaths = logFilesPaths + fullFilePath + "|";
                Directory.CreateDirectory(directoryPath);
                Serialize<List<FileCheckingResult>>(file.CheckingResults, fullFilePath);
            }
            
            Properties.Settings.Default.logFilesPaths = logFilesPaths;
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// Выполняет все проверки, сохраняет их результаты во внешний файл и сохраняет выбор файлов для последующего запуска
        /// </summary>
        private void UpdateAll()
        {
            if (UpdateForSelectedFile)
            {
                Document doc = OpenRvtFile(SelectedFile?.Path);
                Properties.Settings.Default.CurrentFile = SelectedFile?.Path;
                Properties.Settings.Default.Save();
                FileCheckingResult lastResult = SelectedFile?.CheckingResults[SelectedFile.CheckingResults.Count - 1];
                foreach (CheckingTemplate Checking in lastResult.Checkings)
                {
                    UpdateChecking(Checking, doc);
                }
                CloseRvtFile(doc);
            }
            else
            {
                for (int i = 0; i < PreanalysFiles.Count; i++)
                {
                    View.activeFileComboBox.SelectedIndex = i;
                    Document doc = OpenRvtFile(Model.PreanalysFiles[i].Path);
                    Properties.Settings.Default.CurrentFile = Model.PreanalysFiles[i].Path;
                    Properties.Settings.Default.Save();
                    FileCheckingResult lastResult = Model.PreanalysFiles[i].CheckingResults[SelectedFile.CheckingResults.Count - 1];
                    foreach (CheckingTemplate Checking in lastResult.Checkings)
                    {
                        UpdateChecking(Checking, doc);
                    }
                    CloseRvtFile(doc);
                }
            }
            UpdateSettings();
        }

        /// <summary>
        /// Выполняет выбранные проверки, сохраняет их результаты во внешний файл и сохраняет выбор файлов для последующего запуска
        /// </summary>
        private void UpdateSelected()
        {
            if (UpdateForSelectedFile == true)
            {
                TabItem selectedTabItem = View.CheckingsTabControl.SelectedItem as TabItem;
                DataGrid selectedDataGrid = selectedTabItem.Content as DataGrid;
                Document doc = OpenRvtFile(SelectedFile?.Path);
                Properties.Settings.Default.CurrentFile = SelectedFile?.Path;
                Properties.Settings.Default.Save();
                foreach (var item in selectedDataGrid.SelectedItems)
                {
                    CheckingTemplate activeChecking = (CheckingTemplate)item;
                    foreach (CheckingTemplate checking in SelectedFile?.CheckingResults[0].Checkings)
                    {
                        if (checking.Name == activeChecking.Name)
                        {
                            UpdateChecking(checking, doc);
                        }
                    }
                }
                CloseRvtFile(doc);
            }
            else
            {
                foreach (RvtFileInfo file in Model.PreanalysFiles)
                {
                    Document doc = OpenRvtFile(file.Path);
                    Properties.Settings.Default.CurrentFile = file.Path;
                    Properties.Settings.Default.Save();
                    TabItem selectedTabItem = View.CheckingsTabControl.SelectedItem as TabItem;
                    DataGrid selectedDataGrid = selectedTabItem.Content as DataGrid;
                    foreach (var item in selectedDataGrid.SelectedItems)
                    {
                        CheckingTemplate activeChecking = (CheckingTemplate)item;
                        foreach (CheckingTemplate checking in file.CheckingResults[0].Checkings)
                        {
                            if (checking.Name == activeChecking.Name)
                            {
                                UpdateChecking(checking, doc);
                            }
                        }
                    }
                    CloseRvtFile(doc);
                }
            }
            UpdateSettings();
        }

        /// <summary>
        /// Общий для UpdateSelected и UpdateAll метод запуска проверки и обновления результатов
        /// </summary>
        private void UpdateChecking(CheckingTemplate Checking, Document doc)
        {
            Checking.LastRun = System.DateTime.Now.ToString();
            
            Checking.Status = Checking.Run(doc, Checking.ElementCheckingResults);
            Checking.Amount = Checking.ElementCheckingResults.Count.ToString();
            Checking.Created = Checking.ElementCheckingResults.Where(t => t.Status == "Созданная").ToList().Count.ToString();
            Checking.Active = Checking.ElementCheckingResults.Where(t => t.Status == "Активная").ToList().Count.ToString();
            Checking.Corrected = Checking.ElementCheckingResults.Where(t => t.Status == "Исправленная").ToList().Count.ToString();
            Checking.Checked = Checking.ElementCheckingResults.Where(t => t.Status == "Проверенная").ToList().Count.ToString();
            //TextWriterTraceListener tr = new TextWriterTraceListener(System.IO.File.CreateText("C:\\Users\\Евгений\\source\\repos\\etl_audit\\Output.txt"));
            //Debug.Listeners.Add(tr);
            //Debug.WriteLine("Запущена проверка");
        }

        /// <summary>
        /// Обновляет данные последнего запуска для последующего показа в приложении
        /// </summary>
        private void UpdateSettings()
        {
            string lastAnalysFilesString = "";
            string lastAnalysFilesPathsString = "";
            foreach (RvtFileInfo fileToSaveInList in Model.PreanalysFiles)
            {
                lastAnalysFilesString = lastAnalysFilesString + fileToSaveInList.Name + "|";
                lastAnalysFilesPathsString = lastAnalysFilesPathsString + fileToSaveInList.Path + "|";
            }
            Properties.Settings.Default.lastCheckedFiles = lastAnalysFilesString;
            Properties.Settings.Default.lastCheckedFilesPaths = lastAnalysFilesPathsString;
            Properties.Settings.Default.Save();
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
        /// Выгружает отчет о проверках
        /// </summary>
        public void CreateReport()
        {
            ReportsSettings.CreateReport(View, Model.PreanalysFiles);
        }

        /// <summary>
        /// Открывает файл rvt и возвращает его в виде объекта Document
        /// </summary>
        /// <param name="filePath"> Путь к файлу</param>
        /// <returns></returns>
        public static Document OpenRvtFile(string filePath)
        {
            UIApplication uiapp = CommandLauncher.uiapp;
            ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);
            OpenOptions openOptions = new OpenOptions();
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            WorksetConfiguration worksets = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            openOptions.SetOpenWorksetsConfiguration(worksets);
            Autodesk.Revit.ApplicationServices.Application app = CommandLauncher.app;
            Document doc = app.OpenDocumentFile(path, openOptions);
            return doc;
        }
        public static void CloseRvtFile(Document doc)
        {
            if (doc.Title != CommandLauncher.uiapp.ActiveUIDocument.Document.Title)
            {
                doc.Close(false);
            }
        }

        public void CreateParameterChecking()
        {
            ParameterCheckingViewModel parameterViewModel = new ParameterCheckingViewModel(Model);
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

        private RelayCommand _createReportCommand;
        public RelayCommand CreateReportCommand
        {
            get { return _createReportCommand ?? (_createReportCommand = new RelayCommand(CreateReport)); }
        }

        private RelayCommand _createParameterCheckingCommand;
        public RelayCommand CreateParameterCheckingCommand
        {
            get { return _createParameterCheckingCommand ?? (_createParameterCheckingCommand = new RelayCommand(CreateParameterChecking)); }
        }
        #endregion

        #region Вспомогательные классы для хранения данных

        public class ReportSettings
        {
            public ReportFields Fields { get; set; } = new ReportFields();
            public ReportStatuses Statuses { get; set; } = new ReportStatuses();
            public string Type { get; set; }
            public string Format { get; set; }
            public bool WriteForSelectedFile { get; set; }

            /// <summary>
            /// Создает отчеты о проверках
            /// </summary>
            public void CreateReport(StartWindow view, ObservableCollection<RvtFileInfo> data)
            {
                if (Format.ToString() == "Excel")
                {
                    CreateExcel(view, data);
                }
                else if (Format.ToString() == "HTML")
                {

                }
                Properties.Settings.Default.ReportSettings = Serialize<ReportSettings>(this);
                Properties.Settings.Default.Save();
            }
            /// <summary>
            /// Создает результаты проверок в формат excel
            /// </summary>
            private void CreateExcel(StartWindow view, ObservableCollection<RvtFileInfo> data)
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                folderBrowserDialog.Description = "Выберите папку для сохранения отчетов";
                if (Properties.Settings.Default.LastReportPath != "")
                {
                    folderBrowserDialog.SelectedPath = Properties.Settings.Default.LastReportPath;
                }
                folderBrowserDialog.ShowDialog();
                string path = folderBrowserDialog.SelectedPath;
                Properties.Settings.Default.LastReportPath = path;
                Properties.Settings.Default.Save();
                Excel.Application app = new Excel.Application();
                app.DisplayAlerts = false;
                if (path != "" && path != null)
                {
                    if (WriteForSelectedFile == true)
                    {
                        RvtFileInfo file = view.activeFileComboBox.SelectedItem as RvtFileInfo;
                        Excel.Workbook wb = app.Workbooks.Add();
                        wb.Title = file.Name;
                        if (Type.ToString() == "Все тесты")
                        {
                            foreach (CheckingTemplate checking in file.CheckingResults[0].Checkings)
                            {
                                CreateExcelCheckingWorksheet(wb, checking);
                            }
                        }
                        else if (Type.ToString() == "Выбранные тесты")
                        {
                            TabItem selectedTabItem = view.CheckingsTabControl.SelectedItem as TabItem;
                            DataGrid selectedDataGrid = selectedTabItem.Content as DataGrid;
                            foreach (var item in selectedDataGrid.SelectedItems)
                            {
                                CheckingTemplate activeChecking = (CheckingTemplate)item;
                                CreateExcelCheckingWorksheet(wb, activeChecking);
                            }
                        }

                        string fullPath = path + "\\" + file.Name.Replace(".rvt", ".xlsx");
                        while (File.Exists(fullPath))
                        {
                            fullPath = fullPath.Replace(".xlsx", "") + "(1)" + ".xlsx";
                        }
                        wb.SaveAs(fullPath);
                        app.Quit();
                    }
                    else
                    {
                        foreach (RvtFileInfo file in data)
                        {
                            Excel.Workbook wb = app.Workbooks.Add();
                            wb.Title = file.Name;
                            if (Type.ToString() == "Все тесты")
                            {
                                foreach (CheckingTemplate checking in file.CheckingResults[0].Checkings)
                                {
                                    CreateExcelCheckingWorksheet(wb, checking);
                                }
                            }
                            else if (Type.ToString() == "Выбранные тесты")
                            {
                                TabItem selectedTabItem = view.CheckingsTabControl.SelectedItem as TabItem;
                                DataGrid selectedDataGrid = selectedTabItem.Content as DataGrid;
                                foreach (var item in selectedDataGrid.SelectedItems)
                                {
                                    CheckingTemplate activeChecking = (CheckingTemplate)item;
                                    CreateExcelCheckingWorksheet(wb, activeChecking);
                                }
                            }

                            string fullPath = path + "\\" + file.Name.Replace(".rvt", ".xlsx");
                            while (File.Exists(fullPath))
                            {
                                fullPath = fullPath.Replace(".xlsx", "") + "(1)" + ".xlsx";
                            }
                            wb.SaveAs(fullPath);
                            app.Quit();
                        }
                    }
                    
                }
            }
            /// <summary>
            /// Создает лист с результатом проверки для одной проверки одного файла. Метод используется для предотвращения дублирования кода внутри CreateExcel()
            /// </summary>
            private void CreateExcelCheckingWorksheet(Excel.Workbook wb, CheckingTemplate checking)
            {
                Excel.Worksheet sheet = wb.Worksheets.Add();
                if (checking.Name.Length > 29)
                {
                    sheet.Name = checking.Name.Substring(0, 30);
                }
                else
                {
                    sheet.Name = checking.Name;
                }
                if (checking.ResultType == CheckingResultType.ElementsList)
                {
                    ///Создаем булевые маски из настроек отчета и по ним фильтруем список с информацией о файлах
                    bool[] fieldBoolMask = { Fields.Name,
                            Fields.Status,
                            Fields.Id,
                            Fields.Comment,
                            Fields.Time };
                    List<string> rawFields = new List<string>() { "Имя", "Статус", "ID элемента", "Комментарий", "Время обнаружения" };
                    List<string> readyFields = rawFields.Where(x => fieldBoolMask[rawFields.IndexOf(x)] == true).ToList();
                    bool[] statusesBoolMask = { Statuses.Created,
                        Statuses.Active, Statuses.Checked, Statuses.Corrected};
                    List<string> rawStatuses = new List<string>() { "Созданная", "Активная", "Проверенная", "Исправленная" };
                    List<string> readyStatuses = rawStatuses.Where(x => statusesBoolMask[rawStatuses.IndexOf(x)] == true).ToList();
                    List<List<string>> data = new List<List<string>>();
                    CheckingStatusToTextConverter statusConverter = new CheckingStatusToTextConverter();
                    string strStatus = (string)statusConverter.Convert(checking.Status, typeof(string), new object(), new CultureInfo(0x040A, false));
                    data.Add(new List<string>() { strStatus });
                    data.Add(readyFields);
                    foreach (ElementCheckingResult elementResult in checking.ElementCheckingResults.Where(x => readyStatuses.Contains(x.Status)))
                    {
                        List<string> values = new List<string>();
                        for (int i = 0; i < elementResult.GetType().GetTypeInfo().GetProperties().Length; i++)
                        {
                            if (fieldBoolMask[i] == true)
                            {
                                values.Add((string)elementResult.GetType().GetTypeInfo().GetProperties()[i].GetValue(elementResult));
                            }
                        }
                        data.Add(values);
                    }
                    for (int i = 0; i < data.Count; i++)
                    {
                        for (int j = 0; j < data[i].Count; j++)
                        {
                            sheet.Cells[i + 1, j + 1] = data[i][j];
                        }
                    }
                }
                else if (checking.ResultType == CheckingResultType.Message)
                {
                    sheet.Cells[1, 1] = checking.Message;
                }
            }
        }
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

        public enum CheckingResultType
        {
            Message,
            ElementsList 
        }
        
        public enum CheckingStatus
        {
            CheckingSuccessful,
            CheckingFailed,
            NotUpdated,
            NotLaunched
        }

        #endregion
    }
}
