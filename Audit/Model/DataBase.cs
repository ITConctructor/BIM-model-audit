using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Audit.Model
{
    public class DataBase
    {
        public ObservableCollection<RvtFileInfo> PreanalysFiles { get; set; } = new ObservableCollection<RvtFileInfo>();
    }
}
