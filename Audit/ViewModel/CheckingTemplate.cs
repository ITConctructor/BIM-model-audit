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
using Audit.Checkings;

namespace Audit
{
    [XmlInclude(typeof(IfAllRoomsArePlaced)), 
        XmlInclude(typeof(IfLevelNamesAreCorrectChecking)), 
        XmlInclude(typeof(IfLevelsAreNotDuplicated)), 
        XmlInclude(typeof(IfLinkInCorrectWorksetChecking)), 
        XmlInclude(typeof(IfPilesAreMadeOfColumns))]
    public class CheckingTemplate
    {
        public string Dep { get; set; }
        public string Name { get; set; }
        public string LastRun { get; set; }
        public string Amount { get; set; }
        public string Created { get; set; }
        public string Active { get; set; }
        public string Checked { get; set; }
        public string Corrected { get; set; }
        public BindingList<ElementCheckingResult> ElementCheckingResults { get; set; }
        public CheckingTemplate Running { get; set; }
        public virtual Result Run(string filePath)
        {
            return Result.Succeeded;
        }
    }
}
