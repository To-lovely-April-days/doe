using System;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Views;
using Prism.Commands;
using Prism.Ioc;
using Prism.Mvvm;

namespace MaxChemical.Modules.DOE.ViewModels
{
    public class DOEMainViewModel : BindableBase
    {
        private readonly IContainerProvider _container;
        private readonly ILogService _logger;
        private int _selectedTabIndex;
        private string _currentBatchId = "";

        public DOEMainViewModel(IContainerProvider container, ILogService logger)
        {
            _container = container;
            _logger = logger?.ForContext<DOEMainViewModel>() ?? throw new ArgumentNullException(nameof(logger));
            OpenDesignWizardCommand = new DelegateCommand(OpenDesignWizard);
            NavigateToExecutionCommand = new DelegateCommand<string>(NavigateToExecution);
            NavigateToAnalysisCommand = new DelegateCommand<string>(s => NavigateToAnalysis(s));
        }

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set
            {
                if (SetProperty(ref _selectedTabIndex, value))
                {
                    RaisePropertyChanged(nameof(IsTab0));
                    RaisePropertyChanged(nameof(IsTab1));
                    RaisePropertyChanged(nameof(IsTab2));
                    RaisePropertyChanged(nameof(IsTab3));
                    // 移除: RaisePropertyChanged(nameof(IsTab4));
                }
            }
        }

        public bool IsTab0 => SelectedTabIndex == 0;
        public bool IsTab1 => SelectedTabIndex == 1;
        public bool IsTab2 => SelectedTabIndex == 2;
        public bool IsTab3 => SelectedTabIndex == 3;
        //public bool IsTab4 => SelectedTabIndex == 4;

        public string CurrentBatchId { get => _currentBatchId; set => SetProperty(ref _currentBatchId, value); }

        public DelegateCommand OpenDesignWizardCommand { get; }
        public DelegateCommand<string> NavigateToExecutionCommand { get; }
        public DelegateCommand<string> NavigateToAnalysisCommand { get; }

        public event EventHandler<string>? RequestLoadExecution;
        public event EventHandler<string>? RequestLoadAnalysis;
        public event EventHandler? RequestRefreshHistory;

        private void OpenDesignWizard()
        {
            try
            {
                var wizardView = _container.Resolve<DOEDesignWizardView>();
                string? savedBatchId = null;

                if (wizardView.DataContext is DOEDesignWizardViewModel wizardVm)
                    wizardVm.BatchSaved += (s, id) => savedBatchId = id;

                wizardView.ShowDialog();

                if (!string.IsNullOrEmpty(savedBatchId))
                    NavigateToExecution(savedBatchId);
                else
                    RequestRefreshHistory?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex) { _logger.LogError(ex, "打开设计向导失败"); }
        }

        public void NavigateToExecution(string? batchId)
        {
            if (string.IsNullOrEmpty(batchId)) return;
            CurrentBatchId = batchId;
            SelectedTabIndex = 1;
            RequestLoadExecution?.Invoke(this, batchId);
        }

        private void NavigateToAnalysis(string? batchId)
        {
            if (string.IsNullOrEmpty(batchId)) return;
            CurrentBatchId = batchId;
            SelectedTabIndex = 2;  //  指向合并后的「模型分析」
            RequestLoadAnalysis?.Invoke(this, batchId);
        }
    }
}
