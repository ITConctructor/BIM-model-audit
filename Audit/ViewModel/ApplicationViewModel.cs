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
            LoadBackups();
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

        private string _backupsDirectoryPath = CommandLauncher.backupsDirectoryPath;
        public string BackupsDirectoryPath { get => _backupsDirectoryPath; set { SetProperty(ref _backupsDirectoryPath, value); } }

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

        private ObservableCollection<string> _selectedCheckingsNames;
        public ObservableCollection<string> SelectedCheckingsNames
        {
            get { return _selectedCheckingsNames; }
            set
            {
                _selectedCheckingsNames = value;
            }
        }

        private ObservableCollection<CheckingTemplate> _selectedCheckings;
        public ObservableCollection<CheckingTemplate> SelectedCheckings
        {
            get { return _selectedCheckings; }
            set
            {
                _selectedCheckings = value;
            }
        }

        private ReportSettings _reportsSettings = new ReportSettings() { Type = "Все тесты", Format = "Excel" };
        public ReportSettings ReportsSettings
        {
            get => _reportsSettings;
            set { SetProperty(ref _reportsSettings, value); }
        }

        public List<string> ReportTypes { get; set; } = new List<string>() { "Все тесты", "Выбранные тесты" };

        public List<string> ReportFormats { get; set; } = new List<string>() { "Excel" };

        public static List<string> ResultStatuses { get; set; } = new List<string>() { "Созданная", "Активная", "Проверенная", "Исправленная" };

        #endregion

        #region Методы

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
        private bool SetProperty<T>(ref T storage, T value, [CallerMemberName] string propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(storage, value)) return false;
            storage = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        /// <summary>
        /// Считывает результаты предыдущих проверок
        /// </summary>
        private void LoadBackups()
        {
            string[] logFilesPaths = Properties.Settings.Default.backupsPaths.Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string path in logFilesPaths)
            {
                if (File.Exists(path))
                {
                    string serializedFileInfo = System.IO.File.ReadAllText(path);
                    RvtFileInfo file = Deserialize<RvtFileInfo>(serializedFileInfo);
                    foreach (CheckingTemplate checking in file.CheckingResults.Last().Checkings)
                    {
                        checking.Status = CheckingStatus.NotUpdated;
                    }
                    Model.PreanalysFiles.Add(file);
                }
            }
            string serializedReportSettings = Properties.Settings.Default.ReportSettings;
            if (serializedReportSettings != "")
            {
                ReportsSettings = Deserialize<ReportSettings>(serializedReportSettings);
            }
        }

        /// <summary>
        /// Сохраняет результаты проверок для последующей работы приложения
        /// </summary>
        public void SaveBackups()
        {
            string logFilesPaths = "";
            foreach (RvtFileInfo file in Model.PreanalysFiles)
            {
                string logFileName = file.Name;
                string directoryPath = CommandLauncher.backupsDirectoryPath + "\\log_" + logFileName;
                directoryPath = directoryPath.Replace(".rvt", "");
                string fullFilePath = directoryPath + "\\log_" + logFileName.Substring(0, logFileName.LastIndexOf(".")) + "_" + DateTime.Now.ToString().Replace(' ', '_').Replace(':', '.') + ".xml";
                logFilesPaths = logFilesPaths + fullFilePath + "|";
                Directory.CreateDirectory(directoryPath);
                Serialize<RvtFileInfo>(file, fullFilePath);
            }
            Properties.Settings.Default.backupsPaths = logFilesPaths;
            Properties.Settings.Default.Save();
            Properties.Settings.Default.ReportSettings = Serialize<ReportSettings>(ReportsSettings);
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
                FileCheckingResult lastResult = SelectedFile?.CheckingResults.Last();
                foreach (CheckingTemplate Checking in lastResult.Checkings)
                {
                    Checking.UpdateResults(doc);
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
                    FileCheckingResult lastResult = Model.PreanalysFiles[i].CheckingResults.Last();
                    foreach (CheckingTemplate Checking in lastResult.Checkings)
                    {
                        Checking.UpdateResults(doc);
                    }
                    CloseRvtFile(doc);
                }
            }
        }

        /// <summary>
        /// Выполняет выбранные проверки, сохраняет их результаты во внешний файл и сохраняет выбор файлов для последующего запуска
        /// </summary>
        private void UpdateSelected()
        {
            if (UpdateForSelectedFile == true)
            {
                Document doc = OpenRvtFile(SelectedFile?.Path);
                Properties.Settings.Default.CurrentFile = SelectedFile?.Path;
                Properties.Settings.Default.Save();
                foreach (CheckingTemplate checking in SelectedCheckings)
                {
                    checking.UpdateResults(doc);
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
                    foreach (string checkingName in SelectedCheckingsNames)
                    {
                        foreach (CheckingTemplate checking in file.CheckingResults.Last().Checkings)
                        {
                            if (checking.Name == checkingName)
                            {
                                checking.UpdateResults(doc);
                            }
                        }
                    }
                    CloseRvtFile(doc);
                }
            }
        }

        /// <summary>
        /// Выгружает отчет о проверках
        /// </summary>
        public void CreateReport()
        {
            ReportsSettings.CreateReport(this);
        }

        /// <summary>
        /// Создает проверку по параметрам
        /// </summary>
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
