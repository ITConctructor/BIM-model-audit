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
        XmlInclude(typeof(IfLinksArePinned))]
    public class CheckingTemplate
    {
        public BindingList<ElementCheckingResult> ElementCheckingResults { get; set; }
        public string Dep { get; set; }
        public string Name { get; set; }
        public string LastRun { get; set; }
        private string _amount;
        public string Amount 
        { 
            get
            {
                return _amount;
            }
            set
            {
                if (ElementCheckingResults.Count == 0)
                {
                    _amount = value;
                }
                else
                {
                    _amount = ElementCheckingResults.Count.ToString();
                }
            }
        }

        private string _created;
        public string Created 
        { 
            get
            {
                return _created;
            }
            set
            {
                if (ElementCheckingResults.Where(t => t.Status == "Созданная").ToList().Count == 0)
                {
                    _created = value;
                }
                else
                {
                    _created = ElementCheckingResults.Where(t => t.Status == "Созданная").ToList().Count.ToString();
                }
            }
        }

        private string _active;
        public string Active 
        { 
            get
            {
                return _active;
            }
            set
            {
                if (ElementCheckingResults.Where(t => t.Status == "Активная").ToList().Count == 0)
                {
                    _active = value;
                }
                else
                {
                    _active = ElementCheckingResults.Where(t => t.Status == "Активная").ToList().Count.ToString();
                }
            }
        }

        private string _checked;
        public string Checked 
        { 
            get
            {
                return _checked;
            }
            set
            {
                if (ElementCheckingResults.Where(t => t.Status == "Исправленная").ToList().Count == 0)
                {
                    _checked = value;
                }
                else
                {
                    _checked = ElementCheckingResults.Where(t => t.Status == "Исправленная").ToList().Count.ToString();
                }
            }
        }

        private string _corrected;
        public string Corrected 
        { 
            get
            {
                return _corrected;
            }
            set
            {
                if (ElementCheckingResults.Where(t => t.Status == "Проверенная").ToList().Count == 0)
                {
                    _corrected = value;
                }
                else
                {
                    _corrected = ElementCheckingResults.Where(t => t.Status == "Проверенная").ToList().Count.ToString();
                }
            }
        }
        public virtual void Run(Document doc, BindingList<ElementCheckingResult> oldResults)
        {

        }
    }
}
