using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Reflection;
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
using Autodesk.Revit.UI.Events;
using System.Runtime.InteropServices;
using Application = System.Windows.Forms.Application;
using System.IO;
using Audit.Model;

namespace Audit
{
    [Transaction(TransactionMode.Manual)]
    public class CommandLauncher : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            backupsDirectoryPath = commandData.Application.Application.CurrentUsersAddinsDataFolderPath + "\\Audit_Results";
            if (!System.IO.Directory.Exists(backupsDirectoryPath))
            {
                System.IO.Directory.CreateDirectory(backupsDirectoryPath);
            }
            commandData.Application.DialogBoxShowing += ModelOpenDialogHandler;
            uiapp = commandData.Application;
            app = commandData.Application.Application;
            DataBase dataBase = new DataBase();
            ApplicationViewModel viewModel = new ApplicationViewModel(dataBase);
            return Result.Succeeded;
        }
        /// <summary>
        /// Обрабатывает диалоговые окна, возникающие при открытии модели
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void ModelOpenDialogHandler(object sender, DialogBoxShowingEventArgs e)
        {
            switch (e)
            {
                case DialogBoxShowingEventArgs args1:
                    if (e.DialogId == "TaskDialog_Missing_Third_Party_Updaters")
                    {
                        TaskDialogShowingEventArgs args2 = (TaskDialogShowingEventArgs)args1;
                        e.OverrideResult((int)TaskDialogResult.CommandLink1);
                    }
                    else
                    {
                        string id = e.DialogId;
                        e.OverrideResult(1);
                    }
                    break;
                default:
                    return;
            }
        }
        public static Autodesk.Revit.ApplicationServices.Application app { get; private set; }
        public static UIApplication uiapp { get; private set; }
        public static string backupsDirectoryPath = "";
    }
}
