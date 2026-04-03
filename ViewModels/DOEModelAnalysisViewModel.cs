using MaxChemical.Infrastructure.DOE;
using MaxChemical.Infrastructure.Services;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using Newtonsoft.Json;
using OfficeOpenXml;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MaxChemical.Modules.DOE.ViewModels
{
    public class GPRModelTabItem : BindableBase
    {
        public int ModelId { get; set; }
        public string ModelName { get; set; } = "";
        public string FactorSignature { get; set; } = "";
        public bool IsActive { get; set; }
        public bool IsColdStart { get; set; }
        public int DataCount { get; set; }
        public double? RSquared { get; set; }
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (SetProperty(ref _isSelected, value))
                {
                    RaisePropertyChanged(nameof(TabBackground));
                    RaisePropertyChanged(nameof(TabBorderBrush));
                    RaisePropertyChanged(nameof(TabForeground));
                    RaisePropertyChanged(nameof(TabFontWeight));
                }
            }
        }
        public string ShortInfo => IsActive ? $"R² {RSquared:F2}" : IsColdStart ? $"{DataCount}/6" : "未训练";
        public SolidColorBrush StatusDotColor => IsActive ? new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50))
            : IsColdStart ? new SolidColorBrush(Color.FromRgb(0xFF, 0xA0, 0x00))
            : new SolidColorBrush(Color.FromRgb(0xCC, 0xCC, 0xCC));
        public SolidColorBrush TabBackground => IsSelected ? new SolidColorBrush(Color.FromRgb(0xF5, 0xF9, 0xFF)) : Brushes.White;
        public SolidColorBrush TabBorderBrush => IsSelected ? new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)) : new SolidColorBrush(Color.FromRgb(0xDD, 0xDD, 0xDD));
        public SolidColorBrush TabForeground => IsSelected ? new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)) : new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55));
        public string TabFontWeight => IsSelected ? "SemiBold" : "Normal";
    }

    public class GPRModelDetail : BindableBase
    {
        public int ModelId { get; set; }
        public string ModelName { get; set; } = "";
        public string FactorSignature { get; set; } = "";
        public bool IsActive { get; set; }
        public int DataCount { get; set; }
        public double? RSquared { get; set; }
        public double? RMSE { get; set; }
        public DateTime? LastTrainedTime { get; set; }
        public string? EvolutionHistoryJson { get; set; }
        public string StatusText => IsActive ? "已激活" : DataCount > 0 ? $"冷启动 {DataCount}/6" : "未训练";
        public SolidColorBrush StatusColor => IsActive ? new SolidColorBrush(Color.FromRgb(0x2E, 0x7D, 0x32)) : new SolidColorBrush(Color.FromRgb(0xE6, 0x51, 0x00));
        public SolidColorBrush StatusDotColor => IsActive ? new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50)) : new SolidColorBrush(Color.FromRgb(0xFF, 0xA0, 0x00));
        public string StatusHint => IsActive ? "≥6 组达标" : $"需 {6 - DataCount} 组";
        public string DataCountText => DataCount.ToString();
        public string RSquaredText => IsActive ? $"{RSquared:F4}" : "—";
        public string RMSEText => IsActive ? $"{RMSE:F2}" : "—";
        public string LastTrainedText => LastTrainedTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "—";
        public string StatusSummary => IsActive ? $"已激活 · {DataCount}组 · 上次训练 {LastTrainedTime:HH:mm}" : $"冷启动 {DataCount}/6";
    }

    public class OptimalFactorItem
    {
        public string Key { get; set; } = "";
        /// <summary>
        /// ★ 修复: 从 double 改为 object，支持类别因子字符串值
        /// 连续因子存 double (如 135.2)，类别因子存 string (如 "催化剂B")
        /// </summary>
        public object Value { get; set; } = 0.0;
        public double BarWidth { get; set; }

        /// <summary>
        /// 显示用文本: 连续因子保留4位小数，类别因子直接显示标签
        /// </summary>
        public string DisplayValue => Value is double d ? d.ToString("F4") : Value?.ToString() ?? "";

        /// <summary>
        /// 是否为数值类型（连续因子）
        /// </summary>
        public bool IsNumeric => Value is double or int or long or float;

        /// <summary>
        /// 安全获取数值（类别因子返回 0）
        /// </summary>
        public double NumericValue
        {
            get
            {
                if (Value is double d) return d;
                if (Value is int i) return i;
                if (Value is long l) return l;
                if (Value is float f) return f;
                if (double.TryParse(Value?.ToString(), out var parsed)) return parsed;
                return 0.0;
            }
        }
    }

    /// <summary>
    /// OLS 批次选择项
    /// </summary>
    public class OlsBatchItem : BindableBase
    {
        public string BatchId { get; set; } = "";
        public string BatchName { get; set; } = "";
        public string DesignMethod { get; set; } = "";
        public int CompletedCount { get; set; }
        public int FactorCount { get; set; }
        public string DisplayText => $"{BatchName} ({DesignMethod}, {CompletedCount}组)";
        private bool _isSelected;
        public bool IsSelected { get => _isSelected; set => SetProperty(ref _isSelected, value); }
    }
    /// <summary>
    /// 模型项（用于 Model Reduction CheckBox 列表）
    /// </summary>
    public class ModelTermItem : BindableBase
    {
        /// <summary>项名（如 "温度", "温度×压力", "温度²"）</summary>
        public string TermName { get; set; } = "";

        /// <summary>P 值</summary>
        public double PValue { get; set; }

        /// <summary>P 值显示文本</summary>
        public string PValueText => PValue < 0.001 ? "p<0.001" : $"p={PValue:F3}";

        /// <summary>是否显著 (p < 0.05)</summary>
        public bool IsSignificant => PValue < 0.05;

        /// <summary>是否被用户选中保留在模型中</summary>
        private bool _isIncluded = true;
        public bool IsIncluded { get => _isIncluded; set => SetProperty(ref _isIncluded, value); }
    }
    /// <summary>
    /// 类别因子展开方程项（绑定到 UI ItemsControl）
    /// </summary>
    public class CategoryEquationItem
    {
        public string LevelName { get; set; } = "";
        public string Equation { get; set; } = "";
    }
    /// <summary>
    /// Pareto 图中的效应项 — 支持点击切换
    /// </summary>
    public class ParetoTermItem : BindableBase
    {
        public string TermName { get; set; } = "";
        public double AbsT { get; set; }
        public double PValue { get; set; }
        public bool IsSignificant { get; set; }

        private bool _isIncluded = true;
        public bool IsIncluded { get => _isIncluded; set => SetProperty(ref _isIncluded, value); }
    }


    /// <summary>
    /// 模型分析 ViewModel — 方案B: GPR/OLS 双 Tab 布局
    /// GPR Tab: 按模型查看（跨批次累积）
    /// OLS Tab: 按批次独立分析（先选批次再看 ANOVA）
    /// </summary>
    public class DOEModelAnalysisViewModel : BindableBase
    {
        private readonly IGPRModelService _gprService;
        private readonly IDOEAnalysisService _analysisService;
        private readonly IDOERepository _repository;
        private readonly IFlowParameterProvider _paramProvider;
        private readonly IDialogService _dialogService;
        private readonly DOEExportService _exportService;
        private readonly ILogService _logger;
        private readonly IModelRouter _modelRouter;
        private readonly IDesirabilityService _desirabilityService;

        // GPR
        private ObservableCollection<GPRModelTabItem> _models = new();
        private GPRModelDetail _selectedModel = new();
        private bool _isLoading;
        private string _statusMessage = "";
        private List<OptimalFactorItem> _optimalFactors = new();
        private string _optimalResponseText = "—";
        private string _optimalConfidenceText = "";
        private bool _hasOptimalResult;
        private PlotModel? _evolutionPlot;
        private PlotModel? _sensitivityPlot;
        private PlotModel? _actualVsPredictedPlot;
        private PlotModel? _residualPlot;
        private BitmapImage? _surfaceImage;
        private PlotModel? _surfacePlot;
        private List<string> _factorNames = new();
        private string? _surfaceFactor1;
        private string? _surfaceFactor2;
        private string _importStatusText = "";
        private DataTable _fullTrainingData = new();
        private DataTable _pagedTrainingData = new();
        private int _currentPage = 1;
        private int _totalPages = 1;
        private int _totalDataCount;
        private const int PAGE_SIZE = 10;
        private string _currentFlowId = "";
        private string _currentBatchId = "";

        // Tab 切换
        private bool _isGprTabSelected = true;

        // OLS Tab
        private ObservableCollection<OlsBatchItem> _olsBatchItems = new();
        private OlsBatchItem? _selectedOlsBatch;
        private List<string> _responseNames = new();
        private string _selectedResponseName = "";
        private OLSAnalysisResult? _olsResult;
        private string _olsStatusText = "切换到此页后将自动加载可分析的批次列表";

        // Desirability
        private DesirabilityResult? _desirabilityResult;
        // ── 方程展示 ──
        private List<CategoryEquationItem> _categoryEquations = new();
        private bool _hasCategoricalEquations;
        private string _categoricalFactorName = "";

        public List<CategoryEquationItem> CategoryEquations { get => _categoryEquations; set => SetProperty(ref _categoryEquations, value); }
        public bool HasCategoricalEquations { get => _hasCategoricalEquations; set => SetProperty(ref _hasCategoricalEquations, value); }
        public string CategoricalFactorName { get => _categoricalFactorName; set => SetProperty(ref _categoricalFactorName, value); }

        // ── Pareto 图 + Model Reduction ──
        private PlotModel? _olsParetoPlot;
        private List<ParetoTermItem> _paretoTerms = new();
        private string _includedTermsText = "";
        private string _excludedTermsText = "";

        public PlotModel? OlsParetoPlot { get => _olsParetoPlot; set => SetProperty(ref _olsParetoPlot, value); }
        public string IncludedTermsText { get => _includedTermsText; set => SetProperty(ref _includedTermsText, value); }
        public string ExcludedTermsText { get => _excludedTermsText; set => SetProperty(ref _excludedTermsText, value); }

        // ── 残差四合一 ──
        private PlotModel? _normalProbPlot;
        private PlotModel? _residVsFittedPlot;
        private PlotModel? _residHistogramPlot;
        private PlotModel? _residVsOrderPlot;

        public PlotModel? NormalProbPlot { get => _normalProbPlot; set => SetProperty(ref _normalProbPlot, value); }
        public PlotModel? ResidVsFittedPlot { get => _residVsFittedPlot; set => SetProperty(ref _residVsFittedPlot, value); }
        public PlotModel? ResidHistogramPlot { get => _residHistogramPlot; set => SetProperty(ref _residHistogramPlot, value); }
        public PlotModel? ResidVsOrderPlot { get => _residVsOrderPlot; set => SetProperty(ref _residVsOrderPlot, value); }

        // ── ★ v6: 交互式预测刻画器（JMP 风格拖动联动） ──
        private List<PlotModel> _profilerPlots = new();
        private string _profilerCurrentPredicted = "";
        private string _optimalResultText = "";
        private Dictionary<string, object> _profilerCurrentValues = new();  // 当前各因子值
        private ProfilerResult? _lastProfilerData;  // 缓存最近一次 profiler 数据
        private List<string> _profilerFactorOrder = new();  // 因子顺序（与 PlotModel 列表对齐）
        private bool _isProfilerLoaded;

        public List<PlotModel> ProfilerPlots { get => _profilerPlots; set => SetProperty(ref _profilerPlots, value); }
        public string ProfilerCurrentPredicted { get => _profilerCurrentPredicted; set => SetProperty(ref _profilerCurrentPredicted, value); }
        public string OptimalResultText { get => _optimalResultText; set => SetProperty(ref _optimalResultText, value); }
        public bool IsProfilerLoaded { get => _isProfilerLoaded; set => SetProperty(ref _isProfilerLoaded, value); }

        // ── Commands ──
        public DelegateCommand RefitReducedModelCommand { get; }
        public DelegateCommand RestoreFullModelCommand { get; }
        public DelegateCommand LoadProfilerCommand { get; }
        public DelegateCommand FindOptimalMaxCommand { get; }
        public DelegateCommand FindOptimalMinCommand { get; }

        public DOEModelAnalysisViewModel(
            IGPRModelService gprService, IDOEAnalysisService analysisService,
            IDOERepository repository, IFlowParameterProvider paramProvider,
            IDialogService dialogService, DOEExportService exportService,
            IModelRouter modelRouter, IDesirabilityService desirabilityService,
            ILogService logger)
        {
            _gprService = gprService; _analysisService = analysisService;
            _repository = repository; _paramProvider = paramProvider;
            _dialogService = dialogService; _exportService = exportService;
            _modelRouter = modelRouter; _desirabilityService = desirabilityService;
            _logger = logger?.ForContext<DOEModelAnalysisViewModel>() ?? throw new ArgumentNullException(nameof(logger));

            RefreshAllCommand = new DelegateCommand(async () => await RefreshAllAsync());
            RetrainCommand = new DelegateCommand(async () => await RetrainAsync());
            FindOptimalCommand = new DelegateCommand(async () => await FindOptimalAsync());
            DeleteModelCommand = new DelegateCommand(async () => await DeleteModelAsync());
            GenerateSurfaceCommand = new DelegateCommand(async () => await LoadSurfaceAsync(),
                () => !string.IsNullOrEmpty(SurfaceFactor1) && !string.IsNullOrEmpty(SurfaceFactor2) && SurfaceFactor1 != SurfaceFactor2);
            DownloadTemplateCommand = new DelegateCommand(async () => await DownloadTemplateAsync());
            ImportDataCommand = new DelegateCommand(async () => await ImportDataFromExcelAsync());
            PrevPageCommand = new DelegateCommand(() => GoToPage(CurrentPage - 1), () => CurrentPage > 1);
            NextPageCommand = new DelegateCommand(() => GoToPage(CurrentPage + 1), () => CurrentPage < TotalPages);
            SwitchToGprTabCommand = new DelegateCommand(() => SwitchTab(true));
            SwitchToOlsTabCommand = new DelegateCommand(() => SwitchTab(false));
            SelectOlsBatchCommand = new DelegateCommand<OlsBatchItem>(async b => await SelectOlsBatchAsync(b));
            RunDesirabilityCommand = new DelegateCommand(async () => await RunDesirabilityOptimizationAsync());
            RefitReducedModelCommand = new DelegateCommand(async () => await RefitReducedModelAsync());
            RestoreFullModelCommand = new DelegateCommand(async () => await RestoreFullModelAsync());
            LoadProfilerCommand = new DelegateCommand(async () => await LoadPredictionProfilerAsync());
            FindOptimalMaxCommand = new DelegateCommand(async () => await FindOlsOptimalAsync(true));
            FindOptimalMinCommand = new DelegateCommand(async () => await FindOlsOptimalAsync(false));
        }

        // ══════════════ Properties — GPR ══════════════
        public ObservableCollection<GPRModelTabItem> Models { get => _models; set => SetProperty(ref _models, value); }
        public GPRModelDetail SelectedModel { get => _selectedModel; set => SetProperty(ref _selectedModel, value); }
        public bool IsLoading { get => _isLoading; set => SetProperty(ref _isLoading, value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public List<OptimalFactorItem> OptimalFactors { get => _optimalFactors; set => SetProperty(ref _optimalFactors, value); }
        public string OptimalResponseText { get => _optimalResponseText; set => SetProperty(ref _optimalResponseText, value); }
        public string OptimalConfidenceText { get => _optimalConfidenceText; set => SetProperty(ref _optimalConfidenceText, value); }
        public bool HasOptimalResult { get => _hasOptimalResult; set => SetProperty(ref _hasOptimalResult, value); }
        public PlotModel? EvolutionPlot { get => _evolutionPlot; set => SetProperty(ref _evolutionPlot, value); }
        public PlotModel? SensitivityPlot { get => _sensitivityPlot; set => SetProperty(ref _sensitivityPlot, value); }
        public PlotModel? ActualVsPredictedPlot { get => _actualVsPredictedPlot; set => SetProperty(ref _actualVsPredictedPlot, value); }
        public PlotModel? ResidualPlot { get => _residualPlot; set => SetProperty(ref _residualPlot, value); }
        public BitmapImage? SurfaceImage { get => _surfaceImage; set => SetProperty(ref _surfaceImage, value); }
        public PlotModel? SurfacePlot { get => _surfacePlot; set => SetProperty(ref _surfacePlot, value); }
        public List<string> FactorNames { get => _factorNames; set => SetProperty(ref _factorNames, value); }
        public string? SurfaceFactor1 { get => _surfaceFactor1; set { SetProperty(ref _surfaceFactor1, value); GenerateSurfaceCommand.RaiseCanExecuteChanged(); } }
        public string? SurfaceFactor2 { get => _surfaceFactor2; set { SetProperty(ref _surfaceFactor2, value); GenerateSurfaceCommand.RaiseCanExecuteChanged(); } }
        public DataTable PagedTrainingData { get => _pagedTrainingData; set => SetProperty(ref _pagedTrainingData, value); }
        public int CurrentPage { get => _currentPage; set { if (SetProperty(ref _currentPage, value)) { PrevPageCommand.RaiseCanExecuteChanged(); NextPageCommand.RaiseCanExecuteChanged(); } } }
        public int TotalPages { get => _totalPages; set { if (SetProperty(ref _totalPages, value)) { PrevPageCommand.RaiseCanExecuteChanged(); NextPageCommand.RaiseCanExecuteChanged(); } } }
        public int TotalDataCount { get => _totalDataCount; set => SetProperty(ref _totalDataCount, value); }
        public string ImportStatusText { get => _importStatusText; set => SetProperty(ref _importStatusText, value); }

        // ══════════════ Properties — Tab 切换 ══════════════
        public bool IsGprTabSelected
        {
            get => _isGprTabSelected;
            set { if (SetProperty(ref _isGprTabSelected, value)) { RaisePropertyChanged(nameof(IsOlsTabSelected)); RaisePropertyChanged(nameof(GprTabBackground)); RaisePropertyChanged(nameof(GprTabForeground)); RaisePropertyChanged(nameof(OlsTabBackground)); RaisePropertyChanged(nameof(OlsTabForeground)); } }
        }
        public bool IsOlsTabSelected => !_isGprTabSelected;
        public SolidColorBrush GprTabBackground => IsGprTabSelected ? new SolidColorBrush(Color.FromRgb(0x00, 0x78, 0xD4)) : Brushes.White;
        public SolidColorBrush GprTabForeground => IsGprTabSelected ? Brushes.White : new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55));
        public SolidColorBrush OlsTabBackground => IsOlsTabSelected ? new SolidColorBrush(Color.FromRgb(0x2E, 0x86, 0xC1)) : Brushes.White;
        public SolidColorBrush OlsTabForeground => IsOlsTabSelected ? Brushes.White : new SolidColorBrush(Color.FromRgb(0x55, 0x55, 0x55));

        // ══════════════ Properties — OLS Tab ══════════════
        public ObservableCollection<OlsBatchItem> OlsBatchItems { get => _olsBatchItems; set => SetProperty(ref _olsBatchItems, value); }
        public OlsBatchItem? SelectedOlsBatch { get => _selectedOlsBatch; set => SetProperty(ref _selectedOlsBatch, value); }
        public List<string> ResponseNames { get => _responseNames; set => SetProperty(ref _responseNames, value); }
        public string SelectedResponseName
        {
            get => _selectedResponseName;
            set { if (SetProperty(ref _selectedResponseName, value) && IsOlsTabSelected && SelectedOlsBatch != null) _ = RefreshOlsAnalysisAsync(); }
        }
        public OLSAnalysisResult? OlsResult { get => _olsResult; set { SetProperty(ref _olsResult, value); RaisePropertyChanged(nameof(HasOlsResult)); } }
        public string OlsStatusText { get => _olsStatusText; set => SetProperty(ref _olsStatusText, value); }
        public bool HasOlsResult => OlsResult?.ModelSummary != null;
        public DesirabilityResult? DesirabilityResult { get => _desirabilityResult; set => SetProperty(ref _desirabilityResult, value); }

        // ══════════════ Commands ══════════════
        public DelegateCommand RefreshAllCommand { get; }
        public DelegateCommand RetrainCommand { get; }
        public DelegateCommand FindOptimalCommand { get; }
        public DelegateCommand DeleteModelCommand { get; }
        public DelegateCommand GenerateSurfaceCommand { get; }
        public DelegateCommand DownloadTemplateCommand { get; }
        public DelegateCommand ImportDataCommand { get; }
        public DelegateCommand PrevPageCommand { get; }
        public DelegateCommand NextPageCommand { get; }
        public DelegateCommand SwitchToGprTabCommand { get; }
        public DelegateCommand SwitchToOlsTabCommand { get; }
        public DelegateCommand<OlsBatchItem> SelectOlsBatchCommand { get; }
        public DelegateCommand RunDesirabilityCommand { get; }

        // ══════════════ Tab 切换 ══════════════
        private void SwitchTab(bool toGpr)
        {
            IsGprTabSelected = toGpr;
            if (!toGpr) _ = LoadOlsBatchListAsync();
        }

        // ══════════════ 公开加载方法 ══════════════
        public async Task LoadBatchAsync(string batchId)
        {
            _currentBatchId = batchId;
            var batch = await _repository.GetBatchWithDetailsAsync(batchId);
            if (batch == null) return;
            _currentFlowId = batch.FlowId;
            FactorNames = batch.Factors.Select(f => f.FactorName).ToList();
            ResponseNames = batch.Responses.Select(r => r.ResponseName).ToList();
            _selectedResponseName = ResponseNames.FirstOrDefault() ?? "";
            RaisePropertyChanged(nameof(SelectedResponseName));
            if (FactorNames.Count >= 2) { SurfaceFactor1 = FactorNames[0]; SurfaceFactor2 = FactorNames[1]; }
            await LoadModelsAsync();
            await RefreshAllAsync();
        }

        public async Task LoadAsync() { await LoadModelsAsync(); }

        // ══════════════ OLS 批次列表 ══════════════
        private async Task LoadOlsBatchListAsync()
        {
            try
            {
                IsLoading = true;
                OlsStatusText = "正在加载批次列表...";
                string flowId = !string.IsNullOrEmpty(_currentFlowId) ? _currentFlowId : _paramProvider.GetCurrentFlowId() ?? "";
                var allBatches = await _repository.GetBatchesByFlowAsync(flowId);
                var items = new List<OlsBatchItem>();
                foreach (var b in allBatches)
                {
                    var detail = await _repository.GetBatchWithDetailsAsync(b.BatchId);
                    if (detail == null) continue;
                    var completedCount = detail.Runs.Count(r => r.Status == DOERunStatus.Completed);
                    if (completedCount == 0) continue;
                    items.Add(new OlsBatchItem
                    {
                        BatchId = detail.BatchId,
                        BatchName = detail.BatchName,
                        DesignMethod = detail.DesignMethod.ToString(),
                        CompletedCount = completedCount,
                        FactorCount = detail.Factors.Count
                    });
                }
                OlsBatchItems = new ObservableCollection<OlsBatchItem>(items);
                OlsStatusText = items.Count == 0
                    ? "没有已完成实验的批次。请先执行实验后再进行 OLS 分析。"
                    : $"找到 {items.Count} 个可分析的批次，请选择一个进行 OLS 回归分析。";
            }
            catch (Exception ex) { _logger.LogError(ex, "加载 OLS 批次列表失败"); OlsStatusText = $"加载失败: {ex.Message}"; }
            finally { IsLoading = false; }
        }

        // ══════════════ OLS 选择批次 ══════════════
        private async Task SelectOlsBatchAsync(OlsBatchItem? item)
        {
            if (item == null) return;
            foreach (var b in OlsBatchItems) b.IsSelected = (b.BatchId == item.BatchId);
            SelectedOlsBatch = item;
            var batch = await _repository.GetBatchWithDetailsAsync(item.BatchId);
            if (batch?.Responses?.Count > 0)
            {
                ResponseNames = batch.Responses.Select(r => r.ResponseName).ToList();
                _selectedResponseName = ResponseNames.FirstOrDefault() ?? "";
                RaisePropertyChanged(nameof(SelectedResponseName));
            }
            await RefreshOlsAnalysisAsync();
        }

        // ══════════════ OLS 回归分析 ══════════════
        private async Task RefreshOlsAnalysisAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                IsLoading = true;
                OlsStatusText = "正在进行 OLS 回归分析...";
                OlsResult = await _analysisService.FitOlsAsync(SelectedOlsBatch.BatchId, _selectedResponseName, "quadratic");

                if (OlsResult?.ModelSummary != null)
                {
                    OlsStatusText = $"OLS 分析完成: R²={OlsResult.ModelSummary.RSquared:F4}, R²adj={OlsResult.ModelSummary.RSquaredAdj:F4}";

                    // ★ 方程展示
                    UpdateEquationsDisplay(OlsResult);

                    // ★ Pareto 图
                    await LoadOlsParetoAsync();

                    // ★ 残差四合一
                    await LoadResidualDiagnosticsAsync();
                }
                else { OlsStatusText = "OLS 分析返回空结果"; }
            }
            catch (Exception ex) { _logger.LogError(ex, "OLS 分析失败"); OlsStatusText = $"OLS 分析失败: {ex.Message}"; }
            finally { IsLoading = false; }
        }
        // ═══ 新增方法: 解析方程展示 ═══

        private void UpdateEquationsDisplay(OLSAnalysisResult result)
        {
            var eqInfo = result.ModelSummary.Equations;
            if (eqInfo != null && eqInfo.HasCategorical)
            {
                HasCategoricalEquations = true;
                CategoricalFactorName = eqInfo.CategoricalFactor ?? "";
                CategoryEquations = eqInfo.EquationsByLevel
                    .Select(kv => new CategoryEquationItem
                    {
                        LevelName = $"{CategoricalFactorName} = {kv.Key}",
                        Equation = kv.Value
                    }).ToList();
            }
            else
            {
                HasCategoricalEquations = false;
                CategoryEquations = new();
            }
        }
        // ═══ Pareto 图构建（支持点击切换） ═══

        private async Task LoadOlsParetoAsync()
        {
            try
            {
                var json = await _analysisService.GetEffectsParetoAsync(
                    SelectedOlsBatch!.BatchId, _selectedResponseName);
                var data = JsonConvert.DeserializeObject<List<ParetoEffectItem>>(json);
                if (data == null || data.Count == 0) return;

                // 保存项数据（用于点击交互）
                _paretoTerms = data.Select(d => new ParetoTermItem
                {
                    TermName = d.Term,
                    AbsT = d.AbsT,
                    PValue = d.PValue,
                    IsSignificant = d.Significant,
                    IsIncluded = true  // 初始全部保留
                }).ToList();

                RebuildParetoPlot();
                UpdateTermsText();
            }
            catch (Exception ex) { _logger.LogError(ex, "OLS Pareto 图加载失败"); }
        }
        /// <summary>
        /// 根据 _paretoTerms 的 IsIncluded 状态重新构建 OxyPlot 图
        /// Minitab 风格: 水平条形图，Y轴是项名，X轴是标准化效应(|t|值)
        /// 红色虚线 = t 临界值
        /// </summary>
        private void RebuildParetoPlot()
        {
            var sorted = _paretoTerms.OrderByDescending(t => t.AbsT).ToList();

            var model = new PlotModel
            {
                Title = $"标准化效应 Pareto 图 (α = 0.05)",
                TitleFontSize = 13,
                Subtitle = "基于编码空间 [-1, +1]，已消除因子量纲差异，可直接比较各因子的相对重要性",
                SubtitleFontSize = 9,
                SubtitleColor = OxyColor.FromRgb(120, 120, 120),
                PlotMargins = new OxyThickness(80, 40, 20, 40)
            };

            // Y 轴: 项名（CategoryAxis 从下往上排列）
            var catAxis = new CategoryAxis
            {
                Position = AxisPosition.Left,
                GapWidth = 0.3,
                FontSize = 11
            };
            // Minitab 是从上到下按 |t| 降序，OxyPlot CategoryAxis 从下往上，所以要反转
            for (int i = sorted.Count - 1; i >= 0; i--)
                catAxis.Labels.Add(sorted[i].TermName);

            // X 轴: 标准化效应 |t|
            var valAxis = new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = "标准化效应",
                Minimum = 0,
                AbsoluteMinimum = 0,
                FontSize = 11
            };

            model.Axes.Add(catAxis);
            model.Axes.Add(valAxis);

            // 条形系列
            var series = new BarSeries
            {
                StrokeThickness = 0.5,
                StrokeColor = OxyColor.FromRgb(180, 180, 180),
                LabelPlacement = LabelPlacement.Inside,
                LabelFormatString = "{0:F2}"
            };

            // 反转遍历（与 CategoryAxis 标签顺序一致）
            for (int i = sorted.Count - 1; i >= 0; i--)
            {
                var item = sorted[i];
                var color = item.IsIncluded
                    ? OxyColor.FromRgb(100, 149, 237)  // 蓝色 = 保留
                    : OxyColor.FromArgb(80, 180, 180, 180);  // 浅灰半透明 = 已删除
                series.Items.Add(new BarItem { Value = item.AbsT, Color = color });
            }
            model.Series.Add(series);

            // 红色虚线: t 临界值 = t(1-α/2, df_error)
            // ★ v6: 动态计算，从 ANOVA 表的残差行获取 df_error
            double tCrit = 2.08; // 默认值
            if (OlsResult?.AnovaTable != null)
            {
                var residualRow = OlsResult.AnovaTable.FirstOrDefault(r => r.Source == "残差");
                if (residualRow != null && residualRow.DF > 0)
                {
                    // t(1-0.05/2, df) = t(0.975, df)
                    // 用近似公式: 对 df>30 约为 2.0, df=20 约为 2.086, df=10 约为 2.228
                    // 精确计算需要 t 分布逆函数，这里用 Abramowitz & Stegun 近似
                    double df = residualRow.DF;
                    double p = 0.975; // 1 - α/2 where α = 0.05
                    // 近似: t ≈ z + (z³+z)/(4df) + (5z⁵+16z³+3z)/(96df²) where z = Φ⁻¹(p) ≈ 1.96
                    double z = 1.959964;
                    tCrit = z + (z * z * z + z) / (4 * df)
                              + (5 * Math.Pow(z, 5) + 16 * z * z * z + 3 * z) / (96 * df * df);
                    tCrit = Math.Round(tCrit, 2);
                }
            }

            var critLine = new LineAnnotation
            {
                Type = LineAnnotationType.Vertical,
                X = tCrit,
                Color = OxyColors.Red,
                LineStyle = LineStyle.Dash,
                StrokeThickness = 1.5,
                Text = $"{tCrit:F2}",
                TextColor = OxyColors.Red,
                FontSize = 10,
                TextPosition = new DataPoint(tCrit, 1.02),
                TextHorizontalAlignment = HorizontalAlignment.Left
            };
            model.Annotations.Add(critLine);

            OlsParetoPlot = model;
        }

        /// <summary>
        /// Pareto 图点击处理 — 切换条形的保留/删除状态
        /// 从 code-behind 调用
        /// </summary>
        public void ToggleParetoTerm(int categoryIndex)
        {
            // categoryIndex 是从底部开始的，需要映射回 sorted 列表
            var sorted = _paretoTerms.OrderByDescending(t => t.AbsT).ToList();
            int reversedIndex = sorted.Count - 1 - categoryIndex;
            if (reversedIndex < 0 || reversedIndex >= sorted.Count) return;

            var term = sorted[reversedIndex];
            term.IsIncluded = !term.IsIncluded;

            // 同步回 _paretoTerms
            var original = _paretoTerms.FirstOrDefault(t => t.TermName == term.TermName);
            if (original != null) original.IsIncluded = term.IsIncluded;

            RebuildParetoPlot();
            UpdateTermsText();
        }

        private void UpdateTermsText()
        {
            var included = _paretoTerms.Where(t => t.IsIncluded).Select(t => t.TermName);
            var excluded = _paretoTerms.Where(t => !t.IsIncluded).Select(t => t.TermName);
            IncludedTermsText = string.Join(", ", included);
            ExcludedTermsText = excluded.Any() ? string.Join(", ", excluded) : "无";
        }





        // ═══ 新增方法: 重新拟合精简模型 ═══

        private async Task RefitReducedModelAsync()
        {
            var selectedTerms = _paretoTerms
                .Where(t => t.IsIncluded)
                .Select(t => t.TermName)
                .ToList();

            if (selectedTerms.Count == 0)
            {
                _dialogService.ShowError("至少需要保留一个模型项", "提示");
                return;
            }

            try
            {
                IsLoading = true;
                OlsStatusText = "正在拟合精简模型...";

                OlsResult = await _analysisService.FitOlsCustomAsync(
                    SelectedOlsBatch!.BatchId, _selectedResponseName, selectedTerms);

                if (OlsResult?.ModelSummary != null)
                {
                    OlsStatusText = $"精简模型: R²={OlsResult.ModelSummary.RSquared:F4}, R²adj={OlsResult.ModelSummary.RSquaredAdj:F4}, R²pred={OlsResult.ModelSummary.RSquaredPred:F4}";
                    UpdateEquationsDisplay(OlsResult);

                    // ★ 修复 v6: 更新 Pareto 项的选中状态
                    // Pareto 项名是合并后的 (如 "Catalyst", "Catalyst×Temperature")
                    // 系数表项名是展开的 (如 "Catalyst[2]", "Catalyst[3]", "Catalyst[2]×Temperature")
                    // 策略: 把系数表项名"收缩"回合并名（去掉[...]后缀），再做精确匹配
                    var coeffTerms = OlsResult.Coefficients?
                        .Where(c => c.Term != "截距").Select(c => c.Term).ToList() ?? new();

                    // 收缩: "Catalyst[2]×Temperature" → "Catalyst×Temperature"
                    //        "Catalyst[3]" → "Catalyst"
                    //        "Temperature" → "Temperature"（不变）
                    var collapsedTerms = coeffTerms
                        .Select(ct => System.Text.RegularExpressions.Regex.Replace(ct, @"\[[^\]]+\]", ""))
                        .Distinct()
                        .ToHashSet();

                    foreach (var t in _paretoTerms)
                    {
                        t.IsIncluded = collapsedTerms.Contains(t.TermName);
                    }
                    RebuildParetoPlot();
                    UpdateTermsText();

                    // 更新残差四合一
                    await LoadResidualDiagnosticsAsync();
                }
            }
            catch (Exception ex) { _logger.LogError(ex, "精简模型失败"); OlsStatusText = $"精简模型失败: {ex.Message}"; }
            finally { IsLoading = false; }
        }

        private async Task RestoreFullModelAsync()
        {
            foreach (var t in _paretoTerms) t.IsIncluded = true;
            RebuildParetoPlot();
            UpdateTermsText();
            await RefreshOlsAnalysisAsync();
        }

        // ═══ ★ v6: 交互式预测刻画器（JMP 风格拖动联动） ═══

        private async Task LoadPredictionProfilerAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;

            try
            {
                IsLoading = true;
                OlsStatusText = "正在生成预测刻画器...";

                var json = await _analysisService.GetPredictionProfilerAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName);
                var data = JsonConvert.DeserializeObject<ProfilerResult>(json);
                if (data?.Factors == null) return;

                _lastProfilerData = data;
                _profilerFactorOrder = data.Factors.Keys.ToList();

                // 初始化当前值（中心值）
                _profilerCurrentValues.Clear();
                foreach (var kv in data.Factors)
                    _profilerCurrentValues[kv.Key] = kv.Value.CurrentValue ?? 0;

                BuildProfilerPlots(data);
                IsProfilerLoaded = true;
                OlsStatusText = $"预测刻画器已生成 ({data.Factors.Count} 个因子) — 拖动红色竖线可交互调整";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "预测刻画器加载失败");
                OlsStatusText = $"预测刻画器失败: {ex.Message}";
            }
            finally { IsLoading = false; }
        }

        /// <summary>
        /// ★ v6: 从 code-behind 调用 — 用户拖动红色竖线后更新所有图
        /// ★ 关键: 原地更新每个 PlotModel 的曲线数据，不替换列表（避免 UI 闪烁）
        /// </summary>
        public async Task UpdateProfilerValueAsync(string factorName, object newValue)
        {
            if (!IsProfilerLoaded || SelectedOlsBatch == null) return;

            _profilerCurrentValues[factorName] = newValue;

            try
            {
                var fixedValues = new Dictionary<string, object>(_profilerCurrentValues);
                var json = await _analysisService.GetPredictionProfilerAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName, 50, fixedValues);
                var data = JsonConvert.DeserializeObject<ProfilerResult>(json);
                if (data?.Factors == null) return;

                _lastProfilerData = data;
                foreach (var kv in data.Factors)
                    _profilerCurrentValues[kv.Key] = kv.Value.CurrentValue ?? _profilerCurrentValues[kv.Key];

                // ★ 原地更新每个 PlotModel（不替换 ProfilerPlots 列表）
                UpdateProfilerPlotsInPlace(data, factorName);

                ProfilerCurrentPredicted = $"{data.ResponseName}    {data.CurrentPredicted:F4}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "更新预测刻画器失败");
            }
        }

        /// <summary>
        /// 原地更新 PlotModel 数据:
        /// - 被拖动的连续因子: 曲线不变，只更新十字线（因为扫描的就是自己，曲线不受影响）
        /// - 被拖动的类别因子: 需要重建（红色标记点要移到新水平）
        /// - 其他因子: 重建曲线和十字线（因为固定值变了，曲线形状会变）
        /// </summary>
        private void UpdateProfilerPlotsInPlace(ProfilerResult data, string draggedFactor)
        {
            var currentPlots = ProfilerPlots;
            if (currentPlots == null || currentPlots.Count != _profilerFactorOrder.Count)
            {
                BuildProfilerPlots(data);
                return;
            }

            double yMin = double.MaxValue, yMax = double.MinValue;
            foreach (var kv in data.Factors)
            {
                if (kv.Value.Y.Count > 0)
                {
                    yMin = Math.Min(yMin, kv.Value.Y.Min());
                    yMax = Math.Max(yMax, kv.Value.Y.Max());
                }
            }
            double yPad = (yMax - yMin) * 0.1;
            if (yPad < 0.01) yPad = 1.0;

            for (int i = 0; i < _profilerFactorOrder.Count; i++)
            {
                var fname = _profilerFactorOrder[i];
                if (!data.Factors.TryGetValue(fname, out var fdata)) continue;
                if (i >= currentPlots.Count) continue;

                var pm = currentPlots[i];

                // 只有被拖动的连续因子才走"只更新十字线"逻辑
                bool isDraggedContinuous = (fname == draggedFactor && !fdata.IsCategorical);

                if (isDraggedContinuous)
                {
                    UpdateCrosshairOnly(pm, fdata, data.CurrentPredicted);
                }
                else
                {
                    RebuildSinglePlot(pm, fname, fdata, data.CurrentPredicted, yMin - yPad, yMax + yPad, i == 0 ? data.ResponseName : "");
                }
            }
        }

        /// <summary>
        /// 只更新十字线位置（被拖动的图）
        /// </summary>
        private void UpdateCrosshairOnly(PlotModel pm, ProfilerFactorData fdata, double currentPredicted)
        {
            // 更新红色竖线
            var vline = pm.Annotations
                .OfType<OxyPlot.Annotations.LineAnnotation>()
                .FirstOrDefault(a => a.Type == OxyPlot.Annotations.LineAnnotationType.Vertical && a.Color == OxyColors.Red);
            if (vline != null)
            {
                double curX = fdata.CurrentNumericValue;
                vline.X = curX;
                vline.Text = $"{curX:F1}";
            }

            // 更新红色横线
            var hline = pm.Annotations
                .OfType<OxyPlot.Annotations.LineAnnotation>()
                .FirstOrDefault(a => a.Type == OxyPlot.Annotations.LineAnnotationType.Horizontal);
            if (hline != null)
                hline.Y = currentPredicted;

            pm.InvalidatePlot(true);
        }

        /// <summary>
        /// 重建单个图的全部内容（其他图）
        /// </summary>
        private void RebuildSinglePlot(PlotModel pm, string fname, ProfilerFactorData fdata,
            double currentPredicted, double yMin, double yMax, string yTitle)
        {
            pm.Series.Clear();
            pm.Annotations.Clear();
            pm.Axes.Clear();

            // Y 轴（共享范围）
            pm.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                FontSize = 9,
                Minimum = yMin,
                Maximum = yMax,
                Title = yTitle
            });

            if (fdata.IsCategorical)
            {
                var labels = fdata.X_Labels;
                pm.Axes.Add(new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    FontSize = 9,
                    Minimum = -0.5,
                    Maximum = labels.Count - 0.5,
                    MajorStep = 1,
                    MinorStep = 1,
                    LabelFormatter = val =>
                    {
                        int idx = (int)Math.Round(val);
                        return idx >= 0 && idx < labels.Count ? labels[idx] : "";
                    }
                });

                // 折线 + 散点
                var lineSeries = new LineSeries
                {
                    Color = OxyColors.Gray,
                    StrokeThickness = 1.5,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 5,
                    MarkerFill = OxyColors.SteelBlue
                };
                for (int j = 0; j < fdata.Y.Count && j < labels.Count; j++)
                    lineSeries.Points.Add(new DataPoint(j, fdata.Y[j]));
                pm.Series.Add(lineSeries);

                // 当前选中水平
                var curVal = _profilerCurrentValues.TryGetValue(fname, out var cv) ? cv?.ToString() : "";
                int curIdx = labels.IndexOf(curVal ?? "");
                if (curIdx >= 0 && curIdx < fdata.Y.Count)
                {
                    pm.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                    {
                        Type = OxyPlot.Annotations.LineAnnotationType.Vertical,
                        X = curIdx,
                        Color = OxyColors.Red,
                        LineStyle = LineStyle.Dash,
                        StrokeThickness = 1.5
                    });
                    var hl = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 7, MarkerFill = OxyColors.Red };
                    hl.Points.Add(new ScatterPoint(curIdx, fdata.Y[curIdx]));
                    pm.Series.Add(hl);
                }
            }
            else
            {
                pm.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, FontSize = 9 });

                var line = new LineSeries { Color = OxyColors.SteelBlue, StrokeThickness = 2 };
                for (int j = 0; j < fdata.X.Count; j++)
                    line.Points.Add(new DataPoint(fdata.X[j], fdata.Y[j]));
                pm.Series.Add(line);

                double curX = fdata.CurrentNumericValue;
                pm.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                {
                    Type = OxyPlot.Annotations.LineAnnotationType.Vertical,
                    X = curX,
                    Color = OxyColors.Red,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.5,
                    Text = $"{curX:F1}",
                    TextColor = OxyColors.Red,
                    FontSize = 9
                });
            }

            // 水平预测线（所有图都有）
            pm.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
            {
                Type = OxyPlot.Annotations.LineAnnotationType.Horizontal,
                Y = currentPredicted,
                Color = OxyColor.FromArgb(80, 255, 0, 0),
                LineStyle = LineStyle.Dot,
                StrokeThickness = 1
            });

            pm.InvalidatePlot(true);
        }

        /// <summary>
        /// ★ v6: 用户点击类别因子的某个水平
        /// </summary>
        public async Task UpdateProfilerCategoryAsync(string factorName, string levelValue)
        {
            await UpdateProfilerValueAsync(factorName, levelValue);
        }

        /// <summary>
        /// 获取因子在 profiler 图列表中的索引
        /// </summary>
        public string? GetProfilerFactorName(int plotIndex)
        {
            if (plotIndex >= 0 && plotIndex < _profilerFactorOrder.Count)
                return _profilerFactorOrder[plotIndex];
            return null;
        }

        /// <summary>
        /// 获取指定因子是否为类别因子
        /// </summary>
        public bool IsProfilerFactorCategorical(string factorName)
        {
            return _lastProfilerData?.Factors?.TryGetValue(factorName, out var fd) == true && fd.IsCategorical;
        }

        /// <summary>
        /// 获取类别因子的水平列表
        /// </summary>
        public List<string>? GetProfilerCategoryLevels(string factorName)
        {
            if (_lastProfilerData?.Factors?.TryGetValue(factorName, out var fd) == true && fd.IsCategorical)
                return fd.Levels;
            return null;
        }

        /// <summary>
        /// ★ v6: 获取类别因子在指定水平索引处的 Y 值
        /// </summary>
        public double? GetProfilerCategoryYAtIndex(string factorName, int index)
        {
            if (_lastProfilerData?.Factors == null) return null;
            if (!_lastProfilerData.Factors.TryGetValue(factorName, out var fd)) return null;
            if (!fd.IsCategorical || index < 0 || index >= fd.Y.Count) return null;
            return fd.Y[index];
        }

        /// <summary>
        /// 获取连续因子的范围
        /// </summary>
        public (double min, double max) GetProfilerFactorRange(string factorName)
        {
            if (_lastProfilerData?.Factors?.TryGetValue(factorName, out var fd) == true && fd.Range?.Count == 2)
                return (fd.Range[0], fd.Range[1]);
            return (0, 1);
        }

        /// <summary>
        /// ★ v6: 在缓存的剖面曲线上插值，给定 X 返回 Y（用于拖动时 Y 跟随曲线）
        /// </summary>
        public double? GetProfilerYAtX(string factorName, double xValue)
        {
            if (_lastProfilerData?.Factors == null) return null;
            if (!_lastProfilerData.Factors.TryGetValue(factorName, out var fd)) return null;
            if (fd.IsCategorical || fd.X.Count < 2) return null;

            var xs = fd.X;
            var ys = fd.Y;

            // 限制 X 范围
            if (xValue <= xs[0]) return ys[0];
            if (xValue >= xs[xs.Count - 1]) return ys[ys.Count - 1];

            // 线性插值
            for (int i = 0; i < xs.Count - 1; i++)
            {
                if (xValue >= xs[i] && xValue <= xs[i + 1])
                {
                    double t = (xValue - xs[i]) / (xs[i + 1] - xs[i]);
                    return ys[i] + t * (ys[i + 1] - ys[i]);
                }
            }
            return null;
        }

        /// <summary>
        /// 获取全局 Y 范围（所有图共享）
        /// </summary>
        public (double yMin, double yMax) GetProfilerYRange()
        {
            if (_lastProfilerData?.Factors == null) return (0, 1);
            double yMin = double.MaxValue, yMax = double.MinValue;
            foreach (var kv in _lastProfilerData.Factors)
            {
                if (kv.Value.Y.Count > 0)
                {
                    yMin = Math.Min(yMin, kv.Value.Y.Min());
                    yMax = Math.Max(yMax, kv.Value.Y.Max());
                }
            }
            double yPad = (yMax - yMin) * 0.1;
            return (yMin - yPad, yMax + yPad);
        }

        /// <summary>
        /// 构建所有 profiler 图（统一 Y 轴范围，JMP 风格）
        /// 连续因子: 曲线 + 红色竖虚线（当前值）
        /// 类别因子: 散点+折线 + 红色标记（当前选中水平）
        /// </summary>
        private void BuildProfilerPlots(ProfilerResult data)
        {
            ProfilerCurrentPredicted = $"{data.ResponseName}    {data.CurrentPredicted:F4}";

            // 计算全局 Y 范围（所有图共享）
            double yMin = double.MaxValue, yMax = double.MinValue;
            foreach (var kv in data.Factors)
            {
                if (kv.Value.Y.Count > 0)
                {
                    yMin = Math.Min(yMin, kv.Value.Y.Min());
                    yMax = Math.Max(yMax, kv.Value.Y.Max());
                }
            }
            double yPad = (yMax - yMin) * 0.1;
            if (yPad < 0.01) yPad = 1.0;

            var plots = new List<PlotModel>();
            foreach (var fname in _profilerFactorOrder)
            {
                if (!data.Factors.TryGetValue(fname, out var fdata)) continue;

                var pm = new PlotModel
                {
                    Title = fname,
                    TitleFontSize = 11,
                    PlotMargins = new OxyThickness(50, 10, 10, 35)
                };

                // 共享 Y 轴
                pm.Axes.Add(new LinearAxis
                {
                    Position = AxisPosition.Left,
                    FontSize = 9,
                    Minimum = yMin - yPad,
                    Maximum = yMax + yPad,
                    Title = plots.Count == 0 ? data.ResponseName : ""  // 只有第一个图显示 Y 轴标题
                });

                if (fdata.IsCategorical)
                {
                    // ★ 类别因子: 散点+折线（JMP 风格），X 轴用 0,1,2... 映射
                    var labels = fdata.X_Labels;
                    var catBottomAxis = new LinearAxis
                    {
                        Position = AxisPosition.Bottom,
                        FontSize = 9,
                        Minimum = -0.5,
                        Maximum = labels.Count - 0.5,
                        MajorStep = 1,
                        MinorStep = 1,
                        LabelFormatter = val =>
                        {
                            int idx = (int)Math.Round(val);
                            return idx >= 0 && idx < labels.Count ? labels[idx] : "";
                        }
                    };
                    pm.Axes.Add(catBottomAxis);

                    // 折线
                    var lineSeries = new LineSeries
                    {
                        Color = OxyColors.Gray,
                        StrokeThickness = 1.5,
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 5,
                        MarkerFill = OxyColors.SteelBlue
                    };
                    for (int i = 0; i < fdata.Y.Count && i < labels.Count; i++)
                        lineSeries.Points.Add(new DataPoint(i, fdata.Y[i]));
                    pm.Series.Add(lineSeries);

                    // 当前选中水平：红色大点
                    var curVal = _profilerCurrentValues.TryGetValue(fname, out var cv) ? cv?.ToString() : "";
                    int curIdx = labels.IndexOf(curVal ?? "");
                    if (curIdx >= 0 && curIdx < fdata.Y.Count)
                    {
                        // 红色竖虚线
                        pm.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                        {
                            Type = OxyPlot.Annotations.LineAnnotationType.Vertical,
                            X = curIdx,
                            Color = OxyColors.Red,
                            LineStyle = LineStyle.Dash,
                            StrokeThickness = 1.5
                        });
                        // 红色大点
                        var highlight = new ScatterSeries
                        {
                            MarkerType = MarkerType.Circle,
                            MarkerSize = 7,
                            MarkerFill = OxyColors.Red
                        };
                        highlight.Points.Add(new ScatterPoint(curIdx, fdata.Y[curIdx]));
                        pm.Series.Add(highlight);
                    }

                    // 水平预测线
                    pm.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                    {
                        Type = OxyPlot.Annotations.LineAnnotationType.Horizontal,
                        Y = data.CurrentPredicted,
                        Color = OxyColor.FromArgb(80, 255, 0, 0),
                        LineStyle = LineStyle.Dot,
                        StrokeThickness = 1
                    });
                }
                else
                {
                    // 连续因子: 曲线 + 红色竖虚线
                    pm.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, FontSize = 9 });

                    var line = new LineSeries { Color = OxyColors.SteelBlue, StrokeThickness = 2 };
                    for (int i = 0; i < fdata.X.Count; i++)
                        line.Points.Add(new DataPoint(fdata.X[i], fdata.Y[i]));
                    pm.Series.Add(line);

                    // 红色竖虚线（当前值）
                    double curX = fdata.CurrentNumericValue;
                    pm.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                    {
                        Type = OxyPlot.Annotations.LineAnnotationType.Vertical,
                        X = curX,
                        Color = OxyColors.Red,
                        LineStyle = LineStyle.Dash,
                        StrokeThickness = 1.5,
                        Text = $"{curX:F1}",
                        TextColor = OxyColors.Red,
                        FontSize = 9
                    });

                    // 水平预测线
                    pm.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                    {
                        Type = OxyPlot.Annotations.LineAnnotationType.Horizontal,
                        Y = data.CurrentPredicted,
                        Color = OxyColor.FromArgb(80, 255, 0, 0),
                        LineStyle = LineStyle.Dot,
                        StrokeThickness = 1
                    });
                }
                plots.Add(pm);
            }
            ProfilerPlots = plots;
        }

        private async Task FindOlsOptimalAsync(bool maximize)
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;

            try
            {
                IsLoading = true;
                OlsStatusText = maximize ? "正在搜索最大值..." : "正在搜索最小值...";

                var json = await _analysisService.FindOptimalAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName, maximize);
                var data = JsonConvert.DeserializeObject<OptimalResult>(json);

                if (data?.Success == true)
                {
                    var factorsStr = string.Join("\n",
                        data.OptimalFactors.Select(kv => $"  {kv.Key} = {kv.Value}"));
                    OptimalResultText = $"解是{(maximize ? "最大值" : "最小值")}\n" +
                                        $"预测响应: {data.PredictedResponse:F4}\n" +
                                        $"最优条件:\n{factorsStr}\n" +
                                        $"{data.Note}";
                    OlsStatusText = $"优化完成: {_selectedResponseName}={data.PredictedResponse:F4}";

                    // 用最优条件刷新 profiler（如果已加载）
                    if (IsProfilerLoaded)
                    {
                        // 更新 currentValues
                        foreach (var kv in data.OptimalFactors)
                            _profilerCurrentValues[kv.Key] = kv.Value;

                        try
                        {
                            var profilerJson = await _analysisService.GetPredictionProfilerAsync(
                                SelectedOlsBatch.BatchId, _selectedResponseName, 50, data.OptimalFactors);
                            var profData = JsonConvert.DeserializeObject<ProfilerResult>(profilerJson);
                            if (profData?.Factors != null)
                            {
                                _lastProfilerData = profData;
                                BuildProfilerPlots(profData);
                            }
                        }
                        catch { /* profiler 刷新失败不影响优化结果 */ }
                    }
                }
                else
                {
                    OptimalResultText = $"优化失败: {data?.Error ?? "未知错误"}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OLS优化失败");
                OptimalResultText = $"优化失败: {ex.Message}";
            }
            finally { IsLoading = false; }
        }


        // ═══ 残差四合一面板 ═══

        private async Task LoadResidualDiagnosticsAsync()
        {
            try
            {
                var json = await _analysisService.GetResidualDiagnosticsAsync(
                    SelectedOlsBatch!.BatchId, _selectedResponseName);
                var diag = JsonConvert.DeserializeObject<ResidualDiagnostics>(json);
                if (diag == null) return;

                // (1) 正态概率图
                var npModel = new PlotModel { Title = "正态概率图", TitleFontSize = 11, PlotMargins = new OxyThickness(50, 25, 15, 35) };
                npModel.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "理论分位数", FontSize = 10 });
                npModel.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "标准化残差", FontSize = 10 });
                var npScatter = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 3.5, MarkerFill = OxyColors.SteelBlue };
                for (int i = 0; i < diag.NormalProbability.TheoreticalQuantiles.Count; i++)
                    npScatter.Points.Add(new ScatterPoint(diag.NormalProbability.TheoreticalQuantiles[i], diag.NormalProbability.OrderedResiduals[i]));
                npModel.Series.Add(npScatter);
                // 参考线 y=x
                if (diag.NormalProbability.TheoreticalQuantiles.Count > 0)
                {
                    double min = diag.NormalProbability.TheoreticalQuantiles.Min();
                    double max = diag.NormalProbability.TheoreticalQuantiles.Max();
                    var refLine = new LineSeries { Color = OxyColors.Red, LineStyle = LineStyle.Dash, StrokeThickness = 1 };
                    refLine.Points.Add(new DataPoint(min, min));
                    refLine.Points.Add(new DataPoint(max, max));
                    npModel.Series.Add(refLine);
                }
                NormalProbPlot = npModel;

                // (2) 残差 vs 拟合值
                var rvfModel = new PlotModel { Title = "残差 vs 拟合值", TitleFontSize = 11, PlotMargins = new OxyThickness(50, 25, 15, 35) };
                rvfModel.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "拟合值", FontSize = 10 });
                rvfModel.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "标准化残差", FontSize = 10 });
                var rvfScatter = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 3.5, MarkerFill = OxyColors.DarkOrange };
                for (int i = 0; i < diag.ResidualsVsFitted.Fitted.Count; i++)
                    rvfScatter.Points.Add(new ScatterPoint(diag.ResidualsVsFitted.Fitted[i], diag.ResidualsVsFitted.Residuals[i]));
                rvfModel.Series.Add(rvfScatter);
                // 零线
                if (diag.ResidualsVsFitted.Fitted.Count > 0)
                {
                    var zeroLine = new LineSeries { Color = OxyColors.Gray, LineStyle = LineStyle.Dash, StrokeThickness = 1 };
                    zeroLine.Points.Add(new DataPoint(diag.ResidualsVsFitted.Fitted.Min(), 0));
                    zeroLine.Points.Add(new DataPoint(diag.ResidualsVsFitted.Fitted.Max(), 0));
                    rvfModel.Series.Add(zeroLine);
                }
                ResidVsFittedPlot = rvfModel;

                // (3) 残差直方图
                var histModel = new PlotModel { Title = "残差直方图", TitleFontSize = 11, PlotMargins = new OxyThickness(50, 25, 15, 35) };
                histModel.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "标准化残差", FontSize = 10 });
                histModel.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "频率", FontSize = 10, Minimum = 0 });
                var histSeries = new RectangleBarSeries { StrokeThickness = 0.5, StrokeColor = OxyColors.Gray };
                for (int i = 0; i < diag.ResidualsHistogram.Frequencies.Count; i++)
                {
                    double x0 = diag.ResidualsHistogram.BinEdges[i];
                    double x1 = diag.ResidualsHistogram.BinEdges[i + 1];
                    histSeries.Items.Add(new RectangleBarItem
                    {
                        X0 = x0,
                        X1 = x1,
                        Y0 = 0,
                        Y1 = diag.ResidualsHistogram.Frequencies[i],
                        Color = OxyColor.FromRgb(100, 149, 237)
                    });
                }
                histModel.Series.Add(histSeries);
                ResidHistogramPlot = histModel;

                // (4) 残差 vs 观测顺序
                var rvoModel = new PlotModel { Title = "残差 vs 观测顺序", TitleFontSize = 11, PlotMargins = new OxyThickness(50, 25, 15, 35) };
                rvoModel.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "观测顺序", FontSize = 10 });
                rvoModel.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "标准化残差", FontSize = 10 });
                var rvoLine = new LineSeries { Color = OxyColors.SteelBlue, MarkerType = MarkerType.Circle, MarkerSize = 3, MarkerFill = OxyColors.SteelBlue };
                for (int i = 0; i < diag.ResidualsVsOrder.Order.Count; i++)
                    rvoLine.Points.Add(new DataPoint(diag.ResidualsVsOrder.Order[i], diag.ResidualsVsOrder.Residuals[i]));
                rvoModel.Series.Add(rvoLine);
                // 零线
                if (diag.ResidualsVsOrder.Order.Count > 0)
                {
                    var zeroLine = new LineSeries { Color = OxyColors.Gray, LineStyle = LineStyle.Dash, StrokeThickness = 1 };
                    zeroLine.Points.Add(new DataPoint(1, 0));
                    zeroLine.Points.Add(new DataPoint(diag.ResidualsVsOrder.Order.Max(), 0));
                    rvoModel.Series.Add(zeroLine);
                }
                ResidVsOrderPlot = rvoModel;
            }
            catch (Exception ex) { _logger.LogError(ex, "残差诊断图加载失败"); }
        }

        // ══════════════ Desirability ══════════════
        private async Task RunDesirabilityOptimizationAsync()
        {
            var batchId = SelectedOlsBatch?.BatchId ?? _currentBatchId;
            if (string.IsNullOrEmpty(batchId)) { _dialogService.ShowError("请先选择一个批次", "提示"); return; }
            if (ResponseNames.Count < 2) { _dialogService.ShowError("Desirability 优化需要至少 2 个响应变量", "提示"); return; }
            try
            {
                IsLoading = true; StatusMessage = "正在搜索多响应最优因子组合...";
                var configs = await _desirabilityService.LoadConfigAsync(batchId);
                if (configs.Count == 0) { _dialogService.ShowError("请先在设计向导 Step4 中配置各响应变量的优化目标", "缺少配置"); return; }
                var batch = await _repository.GetBatchWithDetailsAsync(batchId);
                if (batch == null) return;
                await _desirabilityService.ConfigureAsync(batchId, configs, batch.Factors);
                DesirabilityResult = await _desirabilityService.OptimizeAsync();
                StatusMessage = DesirabilityResult?.Success == true ? $"多响应优化完成: 综合 D = {DesirabilityResult.CompositeD:F4}" : $"优化失败: {DesirabilityResult?.ErrorMessage ?? "未知错误"}";
            }
            catch (Exception ex) { _logger.LogError(ex, "Desirability 优化失败"); _dialogService.ShowError($"优化失败: {ex.Message}", "错误"); }
            finally { IsLoading = false; }
        }

        // ══════════════ GPR 模型列表 ══════════════
        private async Task LoadModelsAsync()
        {
            try
            {
                IsLoading = true; StatusMessage = "正在加载模型...";
                var states = await _repository.GetAllGPRModelsAsync();
                var tabs = states.Select(s => new GPRModelTabItem { ModelId = s.Id, ModelName = string.IsNullOrEmpty(s.ModelName) ? s.FactorSignature : s.ModelName, FactorSignature = s.FactorSignature, IsActive = s.IsActive, IsColdStart = s.DataCount > 0 && !s.IsActive, DataCount = s.DataCount, RSquared = s.RSquared }).ToList();
                Models = new ObservableCollection<GPRModelTabItem>(tabs);
                if (tabs.Count > 0) SelectModel(tabs[0].ModelId);
                StatusMessage = $"已加载 {tabs.Count} 个模型";
            }
            catch (Exception ex) { _logger.LogError(ex, "加载 GPR 模型列表失败"); StatusMessage = "加载失败"; }
            finally { IsLoading = false; }
        }

        public void SelectModel(int modelId)
        {
            foreach (var m in Models) m.IsSelected = (m.ModelId == modelId);
            _ = LoadModelDetailAsync(modelId);
        }

        private async Task LoadModelDetailAsync(int modelId)
        {
            try
            {
                var state = await _repository.GetGPRModelByIdAsync(modelId);
                if (state == null) return;
                _currentFlowId = state.FlowId;
                SelectedModel = new GPRModelDetail { ModelId = state.Id, ModelName = string.IsNullOrEmpty(state.ModelName) ? state.FactorSignature : state.ModelName, FactorSignature = state.FactorSignature, IsActive = state.IsActive, DataCount = state.DataCount, RSquared = state.RSquared, RMSE = state.RMSE, LastTrainedTime = state.LastTrainedTime, EvolutionHistoryJson = state.EvolutionHistoryJson };
                if (!string.IsNullOrEmpty(state.FactorSignature))
                {
                    FactorNames = state.FactorSignature.Split(',').ToList();
                    if (!string.IsNullOrEmpty(_currentBatchId))
                    {
                        var batch = await _repository.GetBatchWithDetailsAsync(_currentBatchId);
                        if (batch?.Responses?.Count > 0) { ResponseNames = batch.Responses.Select(r => r.ResponseName).ToList(); if (string.IsNullOrEmpty(_selectedResponseName)) { _selectedResponseName = ResponseNames.FirstOrDefault() ?? ""; RaisePropertyChanged(nameof(SelectedResponseName)); } }
                    }
                    else
                    {
                        var batches = await _repository.GetBatchesByFlowAsync(state.FlowId);
                        var latestBatch = batches.FirstOrDefault();
                        if (latestBatch != null) { _currentBatchId = latestBatch.BatchId; var batchDetail = await _repository.GetBatchWithDetailsAsync(_currentBatchId); if (batchDetail?.Responses?.Count > 0) { ResponseNames = batchDetail.Responses.Select(r => r.ResponseName).ToList(); if (string.IsNullOrEmpty(_selectedResponseName)) { _selectedResponseName = ResponseNames.FirstOrDefault() ?? ""; RaisePropertyChanged(nameof(SelectedResponseName)); } } }
                    }
                    if (FactorNames.Count >= 2) { SurfaceFactor1 = FactorNames[0]; SurfaceFactor2 = FactorNames[1]; }
                }
                BuildTrainingDataTable(state.TrainingDataJson);
                BuildEvolutionPlot(state.EvolutionHistoryJson);
                if (state.IsActive) { await LoadGPRAnalysisChartsAsync(state); await FindOptimalAsync(); }
                else { ClearAnalysisCharts(); HasOptimalResult = false; }
            }
            catch (Exception ex) { _logger.LogError(ex, "加载模型详情失败: {ModelId}", modelId); }
        }

        // ══════════════ 训练数据集分页 ══════════════
        private void BuildTrainingDataTable(string? trainingDataJson)
        {
            var dt = new DataTable();
            if (string.IsNullOrEmpty(trainingDataJson)) { _fullTrainingData = dt; TotalDataCount = 0; TotalPages = 1; CurrentPage = 1; PagedTrainingData = dt; return; }
            try
            {
                var records = JsonConvert.DeserializeObject<List<TrainingDataRecord>>(trainingDataJson);
                if (records == null || records.Count == 0) { _fullTrainingData = dt; TotalDataCount = 0; TotalPages = 1; CurrentPage = 1; PagedTrainingData = dt; return; }
                dt.Columns.Add("#", typeof(int));
                var factorNames = records[0].Factors?.Keys.ToList() ?? new List<string>();
                foreach (var name in factorNames) dt.Columns.Add(name, typeof(string));
                dt.Columns.Add("响应值", typeof(string)); dt.Columns.Add("来源", typeof(string));
                dt.Columns.Add("批次名称", typeof(string)); dt.Columns.Add("时间", typeof(string));
                int idx = 1;
                foreach (var rec in records)
                {
                    var row = dt.NewRow(); row["#"] = idx++;
                    if (rec.Factors != null) foreach (var name in factorNames) row[name] = rec.Factors.TryGetValue(name, out var v) ? v.ToString("F2") : "—";
                    row["响应值"] = rec.Response.ToString("F2");
                    row["来源"] = rec.Source switch { "measured" => "实验采集", "imported" => "手动导入", _ => rec.Source ?? "未知" };
                    row["批次名称"] = string.IsNullOrEmpty(rec.BatchName) ? "—" : rec.BatchName;
                    row["时间"] = string.IsNullOrEmpty(rec.Timestamp) ? "—" : rec.Timestamp;
                    dt.Rows.Add(row);
                }
                _fullTrainingData = dt; TotalDataCount = dt.Rows.Count;
                TotalPages = Math.Max(1, (int)Math.Ceiling((double)TotalDataCount / PAGE_SIZE));
                CurrentPage = 1; ApplyPage();
            }
            catch (Exception ex) { _logger.LogError(ex, "构建训练数据表格失败"); _fullTrainingData = new DataTable(); TotalDataCount = 0; TotalPages = 1; CurrentPage = 1; PagedTrainingData = new DataTable(); }
        }
        private void GoToPage(int page) { if (page < 1 || page > TotalPages) return; CurrentPage = page; ApplyPage(); }
        private void ApplyPage()
        {
            if (_fullTrainingData.Rows.Count == 0) { PagedTrainingData = _fullTrainingData; return; }
            var paged = _fullTrainingData.Clone();
            int start = (CurrentPage - 1) * PAGE_SIZE; int end = Math.Min(start + PAGE_SIZE, _fullTrainingData.Rows.Count);
            for (int i = start; i < end; i++) paged.ImportRow(_fullTrainingData.Rows[i]);
            PagedTrainingData = paged;
        }

        // ══════════════ GPR 分析图表 ══════════════
        private async Task RefreshAllAsync() { if (SelectedModel.ModelId <= 0) return; await LoadModelDetailAsync(SelectedModel.ModelId); }
        private async Task LoadGPRAnalysisChartsAsync(GPRModelState state)
        {
            try { IsLoading = true; StatusMessage = "正在加载 GPR 分析图表..."; await EnsureGPRServiceLoadedAsync(state); BuildSensitivityPlot(); await BuildActualVsPredictedPlotAsync(state); await BuildResidualPlotAsync(state); await LoadSurfaceAsync(); StatusMessage = "图表加载完成"; }
            catch (Exception ex) { _logger.LogError(ex, "加载 GPR 分析图表失败"); StatusMessage = $"加载失败: {ex.Message}"; }
            finally { IsLoading = false; }
        }
        private async Task EnsureGPRServiceLoadedAsync(GPRModelState state) { _currentFlowId = state.FlowId; await _gprService.LoadStateAsync(_currentFlowId); }
        private async Task LoadSurfaceAsync()
        {
            if (string.IsNullOrEmpty(SurfaceFactor1) || string.IsNullOrEmpty(SurfaceFactor2) || SurfaceFactor1 == SurfaceFactor2) return;
            try
            {
                byte[] imageBytes = Array.Empty<byte>();
                if (_gprService.IsActive)
                {
                    imageBytes = await _analysisService.GetResponseSurfaceImageFromGPRAsync(SurfaceFactor1, SurfaceFactor2);
                    if (imageBytes.Length > 0)
                    {
                        var bmp = new BitmapImage();
                        using (var ms = new MemoryStream(imageBytes)) { bmp.BeginInit(); bmp.CacheOption = BitmapCacheOption.OnLoad; bmp.StreamSource = ms; bmp.EndInit(); bmp.Freeze(); }
                        System.Windows.Application.Current.Dispatcher.Invoke(() => { SurfaceImage = bmp; });
                    }
                }
                if (imageBytes.Length == 0 && !string.IsNullOrEmpty(_currentBatchId))
                { try { imageBytes = await _analysisService.GetResponseSurfaceImageAsync(_currentBatchId, SurfaceFactor1, SurfaceFactor2); } catch { } }
            }
            catch (Exception ex) { _logger.LogError(ex, "响应曲面加载失败"); }
        }
        private void BuildSensitivityPlot()
        {
            var model = new PlotModel();
            try { var sensitivity = _gprService.GetSensitivity(); if (sensitivity.Count > 0) { var sorted = sensitivity.OrderByDescending(kv => kv.Value).ToList(); var catAxis = new CategoryAxis { Position = AxisPosition.Left }; model.Axes.Add(catAxis); model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Minimum = 0, Title = "敏感性" }); var barSeries = new BarSeries { FillColor = OxyColor.FromRgb(0x00, 0x78, 0xD4) }; foreach (var kv in sorted) { catAxis.Labels.Add(kv.Key); barSeries.Items.Add(new BarItem(kv.Value)); } model.Series.Add(barSeries); } }
            catch (Exception ex) { _logger.LogError(ex, "构建敏感性图失败"); }
            SensitivityPlot = model;
        }
        private async Task BuildActualVsPredictedPlotAsync(GPRModelState state)
        {
            var model = new PlotModel();
            try { var td = ParseTrainingData(state.TrainingDataJson); if (td.Count == 0) { ActualVsPredictedPlot = model; return; } var actuals = td.Select(d => d.Response).ToList(); var preds = _gprService.PredictBatch(td.Select(d => d.Factors).ToList()); var predictions = preds.Select(p => p.Mean).ToList(); double ssRes = 0, ssTot = 0, mean = actuals.Average(); for (int i = 0; i < actuals.Count; i++) { ssRes += Math.Pow(actuals[i] - predictions[i], 2); ssTot += Math.Pow(actuals[i] - mean, 2); } double r2 = ssTot > 1e-12 ? 1.0 - ssRes / ssTot : 0.0; model.Title = $"实际值 vs GPR 预测值 (R² = {r2:F4})"; model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "实际值" }); model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "GPR 预测值" }); var scatter = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 5, MarkerFill = OxyColors.SteelBlue }; for (int i = 0; i < actuals.Count; i++) scatter.Points.Add(new ScatterPoint(actuals[i], predictions[i])); model.Series.Add(scatter); double min = Math.Min(actuals.Min(), predictions.Min()); double max = Math.Max(actuals.Max(), predictions.Max()); var diag = new LineSeries { Color = OxyColors.Red, LineStyle = LineStyle.Dash }; diag.Points.Add(new DataPoint(min, min)); diag.Points.Add(new DataPoint(max, max)); model.Series.Add(diag); }
            catch (Exception ex) { _logger.LogError(ex, "构建实际vs预测图失败"); }
            ActualVsPredictedPlot = model;
        }
        private async Task BuildResidualPlotAsync(GPRModelState state)
        {
            var model = new PlotModel { Title = "残差 vs GPR 拟合值" };
            try { var td = ParseTrainingData(state.TrainingDataJson); if (td.Count == 0) { ResidualPlot = model; return; } var actuals = td.Select(d => d.Response).ToList(); var preds = _gprService.PredictBatch(td.Select(d => d.Factors).ToList()); var fittedValues = preds.Select(p => p.Mean).ToList(); model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "GPR 拟合值" }); model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "残差" }); var scatter = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 4, MarkerFill = OxyColors.DarkOrange }; for (int i = 0; i < actuals.Count; i++) scatter.Points.Add(new ScatterPoint(fittedValues[i], actuals[i] - fittedValues[i])); model.Series.Add(scatter); double fMin = fittedValues.Min(), fMax = fittedValues.Max(); var zero = new LineSeries { Color = OxyColors.Gray, LineStyle = LineStyle.Dash }; zero.Points.Add(new DataPoint(fMin, 0)); zero.Points.Add(new DataPoint(fMax, 0)); model.Series.Add(zero); }
            catch (Exception ex) { _logger.LogError(ex, "构建残差图失败"); }
            ResidualPlot = model;
        }
        private void ClearAnalysisCharts() { SensitivityPlot = new PlotModel(); ActualVsPredictedPlot = new PlotModel(); ResidualPlot = new PlotModel(); SurfaceImage = null; StatusMessage = "模型未激活，需要至少 6 组数据"; }

        // ══════════════ GPR 操作 ══════════════
        private async Task RetrainAsync() { try { IsLoading = true; StatusMessage = "正在重新训练..."; await _gprService.ResetModelAsync(_currentFlowId, keepData: true); await _gprService.TrainModelAsync(); await _gprService.SaveStateAsync(_currentFlowId); await LoadModelsAsync(); } catch (Exception ex) { _logger.LogError(ex, "重新训练失败"); StatusMessage = $"训练失败: {ex.Message}"; } finally { IsLoading = false; } }
        private async Task FindOptimalAsync()
        {
            try { if (!_gprService.IsActive) { HasOptimalResult = false; return; } var optimal = _gprService.FindOptimal(); if (optimal?.OptimalFactors != null && optimal.OptimalFactors.Count > 0) { var maxVal = optimal.OptimalFactors.Values.DefaultIfEmpty(1).Max(); OptimalFactors = optimal.OptimalFactors.Select(kv => new OptimalFactorItem { Key = kv.Key, Value = kv.Value, BarWidth = maxVal > 0 ? kv.Value / maxVal * 180 : 0 }).ToList(); OptimalResponseText = $"{optimal.PredictedResponse:F2}"; OptimalConfidenceText = $"± {optimal.PredictionStd:F2}"; HasOptimalResult = true; } else { HasOptimalResult = false; } }
            catch (Exception ex) { _logger.LogError(ex, "搜索最优失败"); HasOptimalResult = false; }
        }
        private async Task DeleteModelAsync() { try { if (SelectedModel.ModelId <= 0) return; await _repository.DeleteGPRModelAsync(SelectedModel.ModelId); await LoadModelsAsync(); } catch (Exception ex) { _logger.LogError(ex, "删除模型失败"); } }

        // ══════════════ 数据导入 ══════════════
        private async Task DownloadTemplateAsync()
        {
            try { if (FactorNames.Count == 0) { _dialogService.ShowError("无法生成模板：未找到因子信息。", "提示"); return; } var responseNames = _responseNames.Count > 0 ? _responseNames : new List<string> { "Response" }; var modelName = System.Text.RegularExpressions.Regex.Replace(SelectedModel.ModelName ?? "模型", @"[^\w\u4e00-\u9fff\-]", "_"); var defaultFileName = $"GPR数据模板_{modelName}_{DateTime.Now:yyyyMMdd}.xlsx"; var path = _dialogService.ShowSaveFileDialog("Excel 文件 (*.xlsx)|*.xlsx", ".xlsx", defaultFileName); if (string.IsNullOrEmpty(path)) return; await _exportService.GenerateGPRTemplateAsync(FactorNames, responseNames, path); _dialogService.ShowInfo($"数据模板已保存至:\n{path}", "模板已生成"); }
            catch (Exception ex) { _logger.LogError(ex, "生成数据模板失败"); _dialogService.ShowError($"生成模板失败: {ex.Message}", "错误"); }
        }
        private async Task ImportDataFromExcelAsync()
        {
            try { if (FactorNames.Count == 0) { _dialogService.ShowError("无法导入：未找到因子信息。", "提示"); return; } var path = _dialogService.ShowOpenFileDialog("Excel 文件 (*.xlsx)|*.xlsx"); if (string.IsNullOrEmpty(path)) return; IsLoading = true; StatusMessage = "正在读取 Excel 文件..."; var responseNames = _responseNames.Count > 0 ? _responseNames : new List<string> { "Response" }; var primaryResponse = responseNames[0]; var importedData = await Task.Run(() => ReadImportExcel(path, FactorNames, responseNames)); if (importedData.Errors.Count > 0) { _dialogService.ShowError($"校验失败:\n{string.Join("\n", importedData.Errors.Take(5))}", "导入失败"); StatusMessage = "导入失败"; return; } if (importedData.Rows.Count == 0) { _dialogService.ShowError("无有效数据。", "导入失败"); StatusMessage = "导入失败"; return; } var confirm = _dialogService.ShowConfirmation($"共 {importedData.Rows.Count} 行有效数据，确定导入？", "确认导入"); if (!confirm) { StatusMessage = "已取消"; return; } await EnsureGPRServiceReadyAsync(); int fedCount = 0; foreach (var row in importedData.Rows) if (row.Responses.TryGetValue(primaryResponse, out var respVal)) { _gprService.AppendData(row.Factors, respVal, "imported", batchName: ""); fedCount++; } await _gprService.SaveInitialStateAsync(_currentFlowId); ImportStatusText = $"已导入 {fedCount} 条"; StatusMessage = $"导入完成: {fedCount} 条"; await LoadModelsAsync(); }
            catch (Exception ex) { _logger.LogError(ex, "导入数据失败"); _dialogService.ShowError($"导入失败: {ex.Message}", "错误"); }
            finally { IsLoading = false; }
        }
        private async Task EnsureGPRServiceReadyAsync()
        {
            if (!string.IsNullOrEmpty(_currentBatchId)) { var batch = await _repository.GetBatchWithDetailsAsync(_currentBatchId); if (batch?.Factors?.Count > 0) { await _gprService.InitializeModelAsync(_currentFlowId, batch.Factors); return; } }
            await _gprService.LoadStateAsync(_currentFlowId);
        }
        private static ImportResult ReadImportExcel(string path, List<string> factorNames, List<string> responseNames)
        {
            var result = new ImportResult();
            try { using var package = new ExcelPackage(new FileInfo(path)); var ws = package.Workbook.Worksheets.FirstOrDefault(); if (ws == null) { result.Errors.Add("没有工作表"); return result; } var headerMap = new Dictionary<string, int>(); for (int col = 1; col <= ws.Dimension.End.Column; col++) { var header = ws.Cells[1, col].Text?.Trim(); if (!string.IsNullOrEmpty(header)) headerMap[header] = col; } var allRequired = factorNames.Concat(responseNames).ToList(); var missing = allRequired.Where(n => !headerMap.ContainsKey(n)).ToList(); if (missing.Count > 0) { result.Errors.Add($"缺少列: {string.Join(", ", missing)}"); return result; } for (int row = 2; row <= ws.Dimension.End.Row; row++) { var factors = new Dictionary<string, double>(); var responses = new Dictionary<string, double>(); bool valid = true; foreach (var fname in factorNames) { var cellVal = ws.Cells[row, headerMap[fname]].Text?.Trim(); if (!double.TryParse(cellVal, out var val)) { valid = false; break; } factors[fname] = val; } if (!valid) continue; foreach (var rname in responseNames) { var cellVal = ws.Cells[row, headerMap[rname]].Text?.Trim(); if (!double.TryParse(cellVal, out var val)) { valid = false; break; } responses[rname] = val; } if (valid) result.Rows.Add(new ImportRow { Factors = factors, Responses = responses }); } }
            catch (Exception ex) { result.Errors.Add($"读取失败: {ex.Message}"); }
            return result;
        }
        private class ImportResult { public List<ImportRow> Rows { get; set; } = new(); public List<string> Errors { get; set; } = new(); }
        private class ImportRow { public Dictionary<string, double> Factors { get; set; } = new(); public Dictionary<string, double> Responses { get; set; } = new(); }

        // ══════════════ 图表辅助 ══════════════
        private void BuildEvolutionPlot(string? json)
        {
            var model = new PlotModel(); model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "数据量", Minimum = 0 }); model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "R²", Minimum = 0, Maximum = 1.05 });
            if (!string.IsNullOrEmpty(json)) { try { var history = JsonConvert.DeserializeObject<List<EvolutionPoint>>(json); if (history?.Count > 0) { var series = new LineSeries { Color = OxyColor.FromRgb(0x00, 0x78, 0xD4), StrokeThickness = 1.8, MarkerType = MarkerType.Circle, MarkerSize = 2.5, MarkerFill = OxyColor.FromRgb(0x00, 0x78, 0xD4) }; foreach (var pt in history) series.Points.Add(new DataPoint(pt.DataCount, pt.RSquared)); model.Series.Add(series); } } catch { } }
            EvolutionPlot = model;
        }
        private List<TrainingDataPoint> ParseTrainingData(string? json)
        {
            if (string.IsNullOrEmpty(json)) return new();
            try { var records = JsonConvert.DeserializeObject<List<TrainingDataRecord>>(json); return records?.Where(r => r.Factors != null).Select(r => new TrainingDataPoint { Factors = r.Factors!, Response = r.Response }).ToList() ?? new(); }
            catch { return new(); }
        }
        private class EvolutionPoint { [JsonProperty("data_count")] public int DataCount { get; set; } [JsonProperty("r_squared")] public double RSquared { get; set; } }
        private class TrainingDataRecord { [JsonProperty("factors")] public Dictionary<string, double>? Factors { get; set; } [JsonProperty("response")] public double Response { get; set; } [JsonProperty("source")] public string? Source { get; set; } [JsonProperty("batch_name")] public string? BatchName { get; set; } [JsonProperty("timestamp")] public string? Timestamp { get; set; } }
        private class TrainingDataPoint { public Dictionary<string, double> Factors { get; set; } = new(); public double Response { get; set; } }
        // Pareto 效应项（从 Python effects_pareto 返回）
        private class ParetoEffectItem
        {
            [JsonProperty("term")] public string Term { get; set; } = "";
            [JsonProperty("t_value")] public double TValue { get; set; }
            [JsonProperty("abs_t")] public double AbsT { get; set; }
            [JsonProperty("p_value")] public double PValue { get; set; }
            [JsonProperty("significant")] public bool Significant { get; set; }
            [JsonProperty("bonferroni_significant")] public bool? BonferroniSignificant { get; set; }
            [JsonProperty("df")] public int DF { get; set; } = 1;
        }

        // 残差诊断数据（从 Python residual_diagnostics 返回）
        private class ResidualDiagnostics
        {
            [JsonProperty("normal_probability")] public NormalProbData NormalProbability { get; set; } = new();
            [JsonProperty("residuals_vs_fitted")] public ResidVsFittedData ResidualsVsFitted { get; set; } = new();
            [JsonProperty("residuals_vs_order")] public ResidVsOrderData ResidualsVsOrder { get; set; } = new();
            [JsonProperty("residuals_histogram")] public HistogramData ResidualsHistogram { get; set; } = new();
            [JsonProperty("cooks_distance")] public CooksData CooksDistance { get; set; } = new();
        }

        private class NormalProbData
        {
            [JsonProperty("theoretical_quantiles")] public List<double> TheoreticalQuantiles { get; set; } = new();
            [JsonProperty("ordered_residuals")] public List<double> OrderedResiduals { get; set; } = new();
        }

        private class ResidVsFittedData
        {
            [JsonProperty("fitted")] public List<double> Fitted { get; set; } = new();
            [JsonProperty("residuals")] public List<double> Residuals { get; set; } = new();
        }

        private class ResidVsOrderData
        {
            [JsonProperty("order")] public List<int> Order { get; set; } = new();
            [JsonProperty("residuals")] public List<double> Residuals { get; set; } = new();
        }

        private class HistogramData
        {
            [JsonProperty("bin_edges")] public List<double> BinEdges { get; set; } = new();
            [JsonProperty("frequencies")] public List<int> Frequencies { get; set; } = new();
        }

        private class CooksData
        {
            [JsonProperty("observation")] public List<int> Observation { get; set; } = new();
            [JsonProperty("distance")] public List<double> Distance { get; set; } = new();
        }

        // ── ★ 新增 v5: 预测刻画器 + 最优化 反序列化类 ──

        private class ProfilerResult
        {
            [JsonProperty("factors")] public Dictionary<string, ProfilerFactorData>? Factors { get; set; }
            [JsonProperty("current_predicted")] public double CurrentPredicted { get; set; }
            [JsonProperty("response_name")] public string ResponseName { get; set; } = "";
        }

        private class ProfilerFactorData
        {
            [JsonProperty("x")] public List<object> XRaw { get; set; } = new();
            [JsonProperty("y")] public List<double> Y { get; set; } = new();
            [JsonProperty("is_categorical")] public bool IsCategorical { get; set; }
            [JsonProperty("current_value")] public object? CurrentValue { get; set; }
            [JsonProperty("range")] public List<double>? Range { get; set; }
            [JsonProperty("levels")] public List<string>? Levels { get; set; }

            /// <summary>连续因子的 X 值（double 列表）</summary>
            public List<double> X => IsCategorical ? new() : XRaw.Select(v => Convert.ToDouble(v)).ToList();
            /// <summary>类别因子的 X 标签</summary>
            public List<string> X_Labels => IsCategorical ? XRaw.Select(v => v?.ToString() ?? "").ToList() : new();
            /// <summary>连续因子的当前值</summary>
            public double CurrentNumericValue => IsCategorical ? 0 : Convert.ToDouble(CurrentValue ?? 0);
        }

        private class OptimalResult
        {
            [JsonProperty("optimal_factors")] public Dictionary<string, object> OptimalFactors { get; set; } = new();
            [JsonProperty("predicted_response")] public double PredictedResponse { get; set; }
            [JsonProperty("maximize")] public bool Maximize { get; set; }
            [JsonProperty("success")] public bool Success { get; set; }
            [JsonProperty("in_range")] public bool InRange { get; set; }
            [JsonProperty("note")] public string Note { get; set; } = "";
            [JsonProperty("error")] public string? Error { get; set; }
        }
    }


}