using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.ComponentModel;
using System.Xml.Serialization;
using Audit.Model.Checkings;

namespace Audit
{
    [XmlInclude(typeof(IfAllRoomsArePlaced)), 
        XmlInclude(typeof(IfLevelNamesAreCorrectChecking)), 
        XmlInclude(typeof(IfLevelsAreNotDuplicated)), 
        XmlInclude(typeof(IfLinkInCorrectWorksetChecking)), 
        XmlInclude(typeof(IfPilesAreMadeOfColumns)),
        XmlInclude(typeof(IfNoForeignFamilies)),
        XmlInclude(typeof(IfLinksArePinned)),
        XmlInclude(typeof(IfNoCADFiles)),
        XmlInclude(typeof(IfNoRoomsWithoutNames)),
        XmlInclude(typeof(IfNavisworksViewExist)),
        XmlInclude(typeof(IfViewsNamesAreCorrect)),
        XmlInclude(typeof(IfNoSheetsWithoutNames)),
        XmlInclude(typeof(IfARWallsNamesAreCorrect))]
    public class CheckingTemplate
    {
        public BindingList<ElementCheckingResult> ElementCheckingResults { get; set; }
        public ApplicationViewModel.CheckingStatus Status { get; set; }
        public ApplicationViewModel.CheckingResultType ResultType { get; set; }
        public string Message { get; set; }
        public string Dep { get; set; }
        public string Name { get; set; }
        public string LastRun { get; set; }
        public string Amount { get; set; }
        public string Created { get; set; }
        public string Active { get; set; }
        public string Checked { get; set; }
        public string Corrected { get; set; }
        public virtual ApplicationViewModel.CheckingStatus Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {
            return ApplicationViewModel.CheckingStatus.CheckingSuccessful;
        }
        public void UpdateCounts()
        {
            Amount = ElementCheckingResults.Count.ToString();
            Created = ElementCheckingResults.Where(t => t.Status == "Созданная").ToList().Count.ToString();
            Active = ElementCheckingResults.Where(t => t.Status == "Активная").ToList().Count.ToString();
            Corrected = ElementCheckingResults.Where(t => t.Status == "Исправленная").ToList().Count.ToString();
            Checked = ElementCheckingResults.Where(t => t.Status == "Проверенная").ToList().Count.ToString();
        }
    }
}
