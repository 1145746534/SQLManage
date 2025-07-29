using Prism.Regions;
using SQLManage.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQLManage.ViewModels
{
    internal class MainWindowViewModel
    {

        private IRegionManager _regionManager;


        public MainWindowViewModel(IRegionManager regionManager) {
            _regionManager = regionManager;

            _regionManager.RegisterViewWithRegion("ViewRegion", typeof(ReportManagementView));

        }


    }
}
