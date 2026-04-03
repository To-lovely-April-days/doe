using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using Newtonsoft.Json;
using Prism.Mvvm;

namespace MaxChemical.Modules.DOE.Views
{
    /// <summary>
    /// DOE 方案编辑对话框 — 修改因子上下界/水平数，保存后重新生成矩阵
    /// </summary>
    public partial class DOEBatchEditDialog : Window
    {
        private readonly IDOEDesignService _designService;
        private readonly IDOERepository _repository;
        private readonly ILogService _logger;
        private readonly DOEBatch _batch;

        public DOEBatchEditDialog(
            DOEBatch batch,
            IDOEDesignService designService,
            IDOERepository repository,
            ILogService logger)
        {
            InitializeComponent();

            _batch = batch;
            _designService = designService;
            _repository = repository;
            _logger = logger;

            // 构建编辑用的 ViewModel
            var vm = new BatchEditViewModel
            {
                BatchName = batch.BatchName,
                DesignMethod = batch.DesignMethod.ToString(),
                EditableFactors = new ObservableCollection<EditableFactor>(
                    batch.Factors.Select(f => new EditableFactor
                    {
                        FactorName = f.FactorName,
                        LowerBound = f.LowerBound,
                        UpperBound = f.UpperBound,
                        LevelCount = f.LevelCount,
                        SourceNodeId = f.SourceNodeId,
                        SourceParamName = f.SourceParamName,
                        FactorSource = f.FactorSource,
                        FactorType = f.FactorType
                    }))
            };

            vm.UpdateEstimatedRunCount();
            DataContext = vm;
        }

        /// <summary>保存是否成功</summary>
        public bool SavedSuccessfully { get; private set; }

        /// <summary>保存成功后的 BatchId（用于跳转到执行页）</summary>
        public string? SavedBatchId { get; private set; }

        private async void Save_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not BatchEditViewModel vm) return;

            try
            {
                // 1. 校验
                foreach (var f in vm.EditableFactors)
                {
                    if (f.UpperBound <= f.LowerBound)
                    {
                        MessageBox.Show($"因子 '{f.FactorName}' 的上界必须大于下界。", "校验失败",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    if (f.LevelCount < 2 || f.LevelCount > 10)
                    {
                        MessageBox.Show($"因子 '{f.FactorName}' 的水平数必须在 2~10 之间。", "校验失败",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                // 2. 更新因子
                var updatedFactors = vm.EditableFactors.Select((ef, i) => new DOEFactor
                {
                    BatchId = _batch.BatchId,
                    FactorName = ef.FactorName,
                    LowerBound = ef.LowerBound,
                    UpperBound = ef.UpperBound,
                    LevelCount = ef.LevelCount,
                    FactorSource = ef.FactorSource,
                    SourceNodeId = ef.SourceNodeId,
                    SourceParamName = ef.SourceParamName,
                    FactorType = ef.FactorType,
                    SortOrder = i
                }).ToList();

                await _repository.SaveFactorsAsync(_batch.BatchId, updatedFactors);

                // 3. 重新生成矩阵
                DOEDesignMatrix? newMatrix = _batch.DesignMethod switch
                {
                    DOEDesignMethod.FullFactorial => await _designService.GenerateFullFactorialAsync(updatedFactors),
                    DOEDesignMethod.FractionalFactorial =>
                        await _designService.GenerateFractionalFactorialAsync(updatedFactors, GetResolution()),
                    DOEDesignMethod.Taguchi =>
                        await _designService.GenerateTaguchiAsync(updatedFactors, GetTaguchiTable()),
                    _ => await _designService.GenerateFullFactorialAsync(updatedFactors)
                };

                if (newMatrix != null)
                {
                    // 4. 旧的 Pending runs 标记为 Skipped（逻辑删除，可追溯）
                    await MarkPendingRunsAsSkippedAsync(_batch.BatchId);

                    // 5. 插入新的 Runs（RunIndex 从已有最大值之后开始）
                    var allRuns = await _repository.GetRunsAsync(_batch.BatchId);
                    int startIndex = allRuns.Count > 0 ? allRuns.Max(r => r.RunIndex) + 1 : 0;

                    var newRuns = new List<DOERunRecord>();
                    for (int i = 0; i < newMatrix.RunCount; i++)
                    {
                        newRuns.Add(new DOERunRecord
                        {
                            BatchId = _batch.BatchId,
                            RunIndex = startIndex + i,
                            FactorValuesJson = JsonConvert.SerializeObject(newMatrix.Rows[i]),
                            DataSource = DOEDataSource.Measured,
                            Status = DOERunStatus.Pending
                        });
                    }
                    await _repository.SaveRunsAsync(newRuns);
                }

                _logger.LogInformation("方案编辑保存成功: {BatchId}, {RunCount} 组新矩阵",
                    _batch.BatchId, newMatrix?.RunCount ?? 0);

                SavedSuccessfully = true;
                SavedBatchId = _batch.BatchId;

                MessageBox.Show($"方案已更新，重新生成了 {newMatrix?.RunCount ?? 0} 组实验。",
                    "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存方案编辑失败");
                MessageBox.Show($"保存失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private int GetResolution()
        {
            try
            {
                if (!string.IsNullOrEmpty(_batch.DesignConfigJson))
                {
                    var config = JsonConvert.DeserializeAnonymousType(_batch.DesignConfigJson,
                        new { fractionalResolution = 3 });
                    return config?.fractionalResolution ?? 3;
                }
            }
            catch { }
            return 3;
        }

        private string GetTaguchiTable()
        {
            try
            {
                if (!string.IsNullOrEmpty(_batch.DesignConfigJson))
                {
                    var config = JsonConvert.DeserializeAnonymousType(_batch.DesignConfigJson,
                        new { taguchiTable = "auto" });
                    return config?.taguchiTable ?? "auto";
                }
            }
            catch { }
            return "auto";
        }

        /// <summary>
        /// 将所有 Pending 状态的 Runs 标记为 Skipped（逻辑删除，可追溯）
        /// </summary>
        private async Task MarkPendingRunsAsSkippedAsync(string batchId)
        {
            var runs = await _repository.GetRunsAsync(batchId);
            foreach (var run in runs.Where(r => r.Status == DOERunStatus.Pending))
            {
                run.Status = DOERunStatus.Skipped;
                await _repository.UpdateRunAsync(run);
            }
        }

    }

    /// <summary>编辑对话框的 ViewModel</summary>
    public class BatchEditViewModel : BindableBase
    {
        private string _batchName = "";
        private string _designMethod = "";
        private int _estimatedRunCount;

        public string BatchName { get => _batchName; set => SetProperty(ref _batchName, value); }
        public string DesignMethod { get => _designMethod; set => SetProperty(ref _designMethod, value); }
        public int EstimatedRunCount { get => _estimatedRunCount; set => SetProperty(ref _estimatedRunCount, value); }
        public ObservableCollection<EditableFactor> EditableFactors { get; set; } = new();

        public void UpdateEstimatedRunCount()
        {
            if (EditableFactors.Count == 0) { EstimatedRunCount = 0; return; }

            // 全因子: 各因子水平数的乘积
            int count = 1;
            foreach (var f in EditableFactors)
                count *= Math.Max(1, f.LevelCount);
            EstimatedRunCount = count;
        }
    }

    /// <summary>可编辑的因子项</summary>
    public class EditableFactor : BindableBase
    {
        private string _factorName = "";
        private double _lowerBound;
        private double _upperBound;
        private int _levelCount = 3;

        public string FactorName { get => _factorName; set => SetProperty(ref _factorName, value); }

        public double LowerBound
        {
            get => _lowerBound;
            set { if (SetProperty(ref _lowerBound, value)) RaisePropertyChanged(nameof(StepSize)); }
        }

        public double UpperBound
        {
            get => _upperBound;
            set { if (SetProperty(ref _upperBound, value)) RaisePropertyChanged(nameof(StepSize)); }
        }

        public int LevelCount
        {
            get => _levelCount;
            set { if (SetProperty(ref _levelCount, value)) RaisePropertyChanged(nameof(StepSize)); }
        }

        public double StepSize => LevelCount > 1 ? (UpperBound - LowerBound) / (LevelCount - 1) : 0;

        // 保留原始值（编辑时不改这些）
        public MaxChemical.Infrastructure.DOE.FactorSourceType FactorSource { get; set; }
        public string? SourceNodeId { get; set; }
        public string? SourceParamName { get; set; }
        public DOEFactorType FactorType { get; set; } = DOEFactorType.Continuous;
    }
}