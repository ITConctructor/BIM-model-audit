using Audit.Model;
using Audit.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Audit.ApplicationViewModel;

namespace Audit.ViewModel
{
    public class ParameterCheckingViewModel
    {
        DataBase Model { get; set; }
        public ParameterCheckingViewModel(DataBase model)
        {
            view = new ParameterCheckingWindow(this);
            view.Show();
        }
        public ParameterCheckingWindow view { get; set; }

        #region Методы
        private void AddRule()
        {

        }
        #endregion

        #region Команды
        private RelayCommand _addRuleCommand;
        public RelayCommand AddRuleCommand
        {
            get { return _addRuleCommand ?? (_addRuleCommand = new RelayCommand(AddRule)); }
        }
        #endregion

        public enum RuleTypes
        {
            Filled,
            Contains,
            NotContains,
            StartsWith,
            EndsWith
        }
    }
}
