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

namespace Audit
{
    [Transaction(TransactionMode.Manual)]
    public class CommandLauncher : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            commandData.Application.DialogBoxShowing += DialogHandler;
            uiapp = commandData.Application;
            app = commandData.Application.Application;
            StartWindow win = new StartWindow();
            ApplicationViewModel model = new ApplicationViewModel(win);
            //ViewModel = model;
            //win.Closing += BaseWindow_Closing;
            win.Show();
            //commandData.Application.DialogBoxShowing -= DialogHandler;
            return Result.Succeeded;
        }
        private async void DialogHandler(object sender, DialogBoxShowingEventArgs e)
        {
            switch (e)
            {
                case DialogBoxShowingEventArgs args1:
                    if (e.DialogId != "")
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
        public static ApplicationViewModel ViewModel;

        //private void BaseWindow_Closing(object sender, EventArgs e)
        //{
        //    model.WriteLog();
        //}
    }
}
