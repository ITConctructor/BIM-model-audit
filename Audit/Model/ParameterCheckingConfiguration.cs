using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Audit.Model
{
    public class ParameterCheckingConfiguration
    {
        public List<string> Categories = new List<string>();
        public List<string> Parameters = new List<string>();
        public List<ParameterCheckingRule> Rules = new List<ParameterCheckingRule>();
    }
}
