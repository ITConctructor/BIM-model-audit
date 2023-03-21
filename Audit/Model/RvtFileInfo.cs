using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static Audit.ApplicationViewModel;

namespace Audit.Model
{
    public class RvtFileInfo
    {
        public RvtFileInfo(string name, string path)
        {
            Name = name;
            Path = path;
            Type[] checkings = Assembly.GetExecutingAssembly().GetTypes().Where(t => t.Namespace == "Audit.Model.Checkings").ToArray();
            CheckingResults.Add(new FileCheckingResult());
            foreach (Type checking in checkings)
            {
                CheckingTemplate Checking = Activator.CreateInstance(Type.GetType(checking.FullName)) as CheckingTemplate;
                if (Checking != null)
                {
                    Checking.ElementCheckingResults = new BindingList<ElementCheckingResult>();
                    Checking.Status = CheckingStatus.NotLaunched;
                    CheckingResults[CheckingResults.Count-1].Checkings.Add(Checking);
                }
            }
        }
        public string Name { get; set; }
        public string Path { get; set; }
        public List<FileCheckingResult> CheckingResults { get; set; } = new List<FileCheckingResult>();
        public override string ToString()
        {
            return Name;
        }
    }
}
