using Audit.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Audit.Model
{
    public class ParameterCheckingRule
    {
        public ParameterCheckingViewModel.RuleTypes Type { get; set; }
        public List<string> Content { get; set; } = new List<string>();
    }
}
