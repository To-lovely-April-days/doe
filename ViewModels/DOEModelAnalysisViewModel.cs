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
using OxyPlot.Legends;
using OxyPlot.Series;
using Prism.Commands;
using Prism.Mvvm;
using Python.Runtime;
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
        // ── ★ v7: 不可估计项警告 ──
        private string _inestimableWarning = "";
        private bool _hasInestimableWarning;

        public string InestimableWarning { get => _inestimableWarning; set => SetProperty(ref _inestimableWarning, value); }
        public bool HasInestimableWarning { get => _hasInestimableWarning; set => SetProperty(ref _hasInestimableWarning, value); }

        // ── ★ v7: OLS 响应曲面 ──
        private BitmapImage? _olsSurfaceImage;
        private PlotModel? _olsContourPlot;
        private List<string> _olsContinuousFactors = new();
        private string? _olsSurfaceFactor1;
        private string? _olsSurfaceFactor2;

        public BitmapImage? OlsSurfaceImage { get => _olsSurfaceImage; set => SetProperty(ref _olsSurfaceImage, value); }
        public PlotModel? OlsContourPlot { get => _olsContourPlot; set => SetProperty(ref _olsContourPlot, value); }
        public List<string> OlsContinuousFactors { get => _olsContinuousFactors; set => SetProperty(ref _olsContinuousFactors, value); }
        public string? OlsSurfaceFactor1
        {
            get => _olsSurfaceFactor1;
            set { SetProperty(ref _olsSurfaceFactor1, value); GenerateOlsSurfaceCommand?.RaiseCanExecuteChanged(); }
        }
        public string? OlsSurfaceFactor2
        {
            get => _olsSurfaceFactor2;
            set { SetProperty(ref _olsSurfaceFactor2, value); GenerateOlsSurfaceCommand?.RaiseCanExecuteChanged(); }
        }
        // ════════════ 1. 新增字段和属性（在 OLS 响应曲面属性区域之后添加） ════════════

        // ── ★ v8: 异常点分析 ──
        private List<OutlierItem> _outlierItems = new();
        private bool _hasOutliers;
        private string _outlierSummaryText = "";

        public List<OutlierItem> OutlierItems { get => _outlierItems; set => SetProperty(ref _outlierItems, value); }
        public bool HasOutliers { get => _hasOutliers; set => SetProperty(ref _hasOutliers, value); }
        public string OutlierSummaryText { get => _outlierSummaryText; set => SetProperty(ref _outlierSummaryText, value); }

        // ── ★ v8: 主效应图 + 交互效应图 ──
        private PlotModel? _olsMainEffectsPlot;
        private PlotModel? _olsInteractionPlot;

        public PlotModel? OlsMainEffectsPlot { get => _olsMainEffectsPlot; set => SetProperty(ref _olsMainEffectsPlot, value); }
        public PlotModel? OlsInteractionPlot { get => _olsInteractionPlot; set => SetProperty(ref _olsInteractionPlot, value); }

        // ── ★ v8: Tukey HSD ──
        private List<TukeyComparisonItem> _tukeyComparisons = new();
        private string _tukeyFactorName = "";
        private bool _hasTukeyResult;
        private Dictionary<string, double> _tukeyGroupMeans = new();

        // ── ★ v9: Box-Cox ──
        private string _boxCoxRecommendation = "";
        private bool _hasBoxCoxResult;
        private double _boxCoxLambda;
        private string _boxCoxTransformName = "";
        private double _boxCoxOrigR2;
        private double _boxCoxTransR2;
        private bool _boxCoxRecommend;
        private PlotModel? _boxCoxProfilePlot;
        // ── ★ v13: 编码/未编码切换 ──
        private bool _isCodedEquation = true;
        public bool IsCodedEquation
        {
            get => _isCodedEquation;
            set
            {
                if (SetProperty(ref _isCodedEquation, value) && OlsResult != null)
                {
                    UpdateEquationsDisplay(OlsResult);
                    RaisePropertyChanged(nameof(EquationModeText));
                }
            }
        }
        public string EquationModeText => IsCodedEquation ? "编码单位" : "未编码单位";
        public string BoxCoxRecommendation { get => _boxCoxRecommendation; set => SetProperty(ref _boxCoxRecommendation, value); }
        public bool HasBoxCoxResult { get => _hasBoxCoxResult; set => SetProperty(ref _hasBoxCoxResult, value); }
        public double BoxCoxLambda { get => _boxCoxLambda; set => SetProperty(ref _boxCoxLambda, value); }
        public string BoxCoxTransformName { get => _boxCoxTransformName; set => SetProperty(ref _boxCoxTransformName, value); }
        public double BoxCoxOrigR2 { get => _boxCoxOrigR2; set => SetProperty(ref _boxCoxOrigR2, value); }
        public double BoxCoxTransR2 { get => _boxCoxTransR2; set => SetProperty(ref _boxCoxTransR2, value); }
        public bool BoxCoxRecommend { get => _boxCoxRecommend; set => SetProperty(ref _boxCoxRecommend, value); }
        public PlotModel? BoxCoxProfilePlot { get => _boxCoxProfilePlot; set => SetProperty(ref _boxCoxProfilePlot, value); }
        public List<TukeyComparisonItem> TukeyComparisons { get => _tukeyComparisons; set => SetProperty(ref _tukeyComparisons, value); }
        public string TukeyFactorName { get => _tukeyFactorName; set => SetProperty(ref _tukeyFactorName, value); }
        public bool HasTukeyResult { get => _hasTukeyResult; set => SetProperty(ref _hasTukeyResult, value); }
        private double _paretoLogWorthCrit = 1.301;
        private double _paretoAlpha = 0.05;
        private string _currentEquationText = "";
        public string CurrentEquationText { get => _currentEquationText; set => SetProperty(ref _currentEquationText, value); }

        // ── OLS 直接导入 (不依赖批次) ──
        private bool _isDirectImportMode;
        private string _directImportFilePath = "";
        private List<string> _directImportFactorNames = new();
        private List<string> _directImportResponseNames = new();
        private Dictionary<string, string> _directImportFactorTypes = new();
        private List<Dictionary<string, object>> _directImportFactorsData = new();
        private Dictionary<string, List<double>> _directImportResponsesData = new();
        /// <summary>是否处于直接导入模式（非批次模式）</summary>
        public bool IsDirectImportMode
        {
            get => _isDirectImportMode;
            set { SetProperty(ref _isDirectImportMode, value); RaisePropertyChanged(nameof(ShowDirectImportPanel)); }
        }

        /// <summary>直接导入的文件路径</summary>
        public string DirectImportFilePath
        {
            get => _directImportFilePath;
            set => SetProperty(ref _directImportFilePath, value);
        }

        /// <summary>是否显示直接导入面板（始终显示）</summary>
        public bool ShowDirectImportPanel => true;
        // ── Commands ──
        public DelegateCommand RefitReducedModelCommand { get; }
        public DelegateCommand RestoreFullModelCommand { get; }
        public DelegateCommand LoadProfilerCommand { get; }
        public DelegateCommand FindOptimalMaxCommand { get; }
        public DelegateCommand FindOptimalMinCommand { get; }
        // 在构造函数末尾添加:
        public DelegateCommand GenerateOlsSurfaceCommand { get; }
        public DelegateCommand ExcludeOutliersCommand { get; }
        public DelegateCommand RestoreAllDataCommand { get; }
        public DelegateCommand ApplyBoxCoxCommand { get; }
        public DelegateCommand ExportReportCommand { get; }
        // 声明:
        public DelegateCommand ToggleEquationCodingCommand { get; }
        public DelegateCommand OlsImportExcelCommand { get; }
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
            GenerateOlsSurfaceCommand = new DelegateCommand(
                async () => await LoadOlsSurfaceAsync(),
                () => !string.IsNullOrEmpty(OlsSurfaceFactor1) && !string.IsNullOrEmpty(OlsSurfaceFactor2)
                       && OlsSurfaceFactor1 != OlsSurfaceFactor2);
            // 构造函数中初始化（添加到末尾）:
            ExcludeOutliersCommand = new DelegateCommand(async () => await ExcludeOutliersAsync());
            RestoreAllDataCommand = new DelegateCommand(async () => await RestoreAllDataAsync());
            ApplyBoxCoxCommand = new DelegateCommand(async () => await ApplyBoxCoxAsync());
            ExportReportCommand = new DelegateCommand(async () => await ExportOlsReportAsync());
            ToggleEquationCodingCommand = new DelegateCommand(() => IsCodedEquation = !IsCodedEquation);
            OlsImportExcelCommand = new DelegateCommand(async () => await OlsImportExcelAsync());
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
            set
            {
                if (SetProperty(ref _selectedResponseName, value) && IsOlsTabSelected)
                {
                    if (IsDirectImportMode)
                        _ = RunDirectOlsAnalysisAsync();  // ★ 直接导入模式下切换响应
                    else if (SelectedOlsBatch != null)
                        _ = RefreshOlsAnalysisAsync();
                }
            }
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
        private async Task LoadOlsSurfaceAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;
            if (string.IsNullOrEmpty(OlsSurfaceFactor1) || string.IsNullOrEmpty(OlsSurfaceFactor2)
                || OlsSurfaceFactor1 == OlsSurfaceFactor2) return;

            try
            {
                // 1. 加载 3D 曲面图（matplotlib 生成的 PNG）
                var imageBytes = await _analysisService.GetOlsResponseSurfaceImageAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName,
                    OlsSurfaceFactor1, OlsSurfaceFactor2);

                if (imageBytes.Length > 0)
                {
                    var bmp = new BitmapImage();
                    using (var ms = new MemoryStream(imageBytes))
                    {
                        bmp.BeginInit();
                        bmp.CacheOption = BitmapCacheOption.OnLoad;
                        bmp.StreamSource = ms;
                        bmp.EndInit();
                        bmp.Freeze();
                    }
                    System.Windows.Application.Current.Dispatcher.Invoke(() => { OlsSurfaceImage = bmp; });
                }
                else
                {
                    OlsSurfaceImage = null;
                }

                // 2. 加载等高线数据（OxyPlot 渲染）
                var contourJson = await _analysisService.GetOlsContourDataAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName,
                    OlsSurfaceFactor1, OlsSurfaceFactor2);

                var contourData = JsonConvert.DeserializeObject<SurfaceData>(contourJson);
                if (contourData != null && contourData.Z != null)
                {
                    BuildContourPlot(contourData);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OLS 响应曲面加载失败");
            }
        }
        /// <summary>
        /// ★ v7: 用 OxyPlot HeatMapSeries 渲染等高线图
        /// </summary>
        private void BuildContourPlot(SurfaceData data)
        {
            var model = new PlotModel
            {
                Title = $"OLS 等高线: {data.XLabel} × {data.YLabel}",
                TitleFontSize = 12,
                PlotMargins = new OxyThickness(60, 30, 80, 40)
            };

            int nx = data.X.Count;
            int ny = data.Y.Count;
            if (nx < 2 || ny < 2) { OlsContourPlot = model; return; }

            // 使用 HeatMapSeries
            var heatMap = new HeatMapSeries
            {
                X0 = data.X[0],
                X1 = data.X[nx - 1],
                Y0 = data.Y[0],
                Y1 = data.Y[ny - 1],
                Interpolate = true,
                RenderMethod = HeatMapRenderMethod.Bitmap,
                Data = new double[nx, ny]
            };

            for (int j = 0; j < ny; j++)
                for (int i = 0; i < nx; i++)
                    heatMap.Data[i, j] = data.Z[j][i];

            model.Series.Add(heatMap);

            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = data.XLabel,
                FontSize = 10
            });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = data.YLabel,
                FontSize = 10
            });

            // 颜色轴
            model.Axes.Add(new LinearColorAxis
            {
                Position = AxisPosition.Right,
                Palette = OxyPalettes.Viridis(200),
                Title = _selectedResponseName,
                FontSize = 9
            });

            OlsContourPlot = model;
        }
        public async Task LoadAsync() { await LoadModelsAsync(); }

        // ═══ ★ v8: 异常点分析面板 ═══

        private async Task LoadOutlierAnalysisAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await _analysisService.GetOutlierAnalysisAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName);
                var data = JsonConvert.DeserializeObject<OutlierAnalysisResult>(json);
                if (data == null) return;

                OutlierItems = data.Outliers.Select(o => new OutlierItem
                {
                    Index = o.Index,
                    Actual = o.Actual,
                    Predicted = o.Predicted,
                    Residual = o.Residual,
                    StdResidual = o.StdResidual,
                    CooksD = o.CooksD,
                    Leverage = o.Leverage,
                    Reasons = string.Join("; ", o.Reasons),
                    IsExcluded = false
                }).ToList();

                HasOutliers = data.OutlierCount > 0;
                OutlierSummaryText = data.OutlierCount > 0
                    ? $"检测到 {data.OutlierCount} 个异常点（共 {data.TotalObservations} 个观测）。阈值: Cook's D > {data.Thresholds.CooksD:F3}, |标准化残差| > {data.Thresholds.StdResidual}, 杠杆值 > {data.Thresholds.Leverage:F3}"
                    : $"未检测到异常点（共 {data.TotalObservations} 个观测）";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "异常点分析失败");
                HasOutliers = false;
            }
        }

        private async Task ExcludeOutliersAsync()
        {
            var excluded = OutlierItems.Where(o => o.IsExcluded).Select(o => o.Index).ToList();
            if (excluded.Count == 0)
            {
                _dialogService.ShowError("请先勾选要排除的异常点", "提示");
                return;
            }

            try
            {
                IsLoading = true;
                OlsStatusText = $"正在排除 {excluded.Count} 个异常点后重新拟合...";

                OlsResult = await _analysisService.RefitExcludingAsync(
                    SelectedOlsBatch!.BatchId, _selectedResponseName, excluded);

                if (OlsResult?.ModelSummary != null)
                {
                    OlsStatusText = $"排除 {excluded.Count} 点后: R²={OlsResult.ModelSummary.RSquared:F4}, R²adj={OlsResult.ModelSummary.RSquaredAdj:F4}";
                    UpdateEquationsDisplay(OlsResult);
                    await LoadOlsParetoAsync();
                    await LoadResidualDiagnosticsAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "排除异常点重拟合失败");
                OlsStatusText = $"重拟合失败: {ex.Message}";
            }
            finally { IsLoading = false; }
        }

        private async Task RestoreAllDataAsync()
        {
            // 恢复完整数据，重新运行 OLS 分析
            await RefreshOlsAnalysisAsync();
        }

        // ═══ ★ v8: 主效应图 ═══

        private async Task LoadOlsMainEffectsAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await _analysisService.GetMainEffectsForResponseAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName);
                var data = JsonConvert.DeserializeObject<Dictionary<string, List<MainEffectPoint>>>(json);
                if (data == null || data.Count == 0) return;

                var model = new PlotModel
                {
                    Title = "主效应图",
                    TitleFontSize = 12,
                    PlotMargins = new OxyThickness(50, 25, 15, 35)
                };
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = _selectedResponseName, FontSize = 10 });

                // 所有因子共用 X 轴（用索引，标签通过 tooltip 显示）
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "因子水平", FontSize = 10 });

                var colors = new[] { OxyColors.Blue, OxyColors.Red, OxyColors.Green, OxyColors.Orange,
                                     OxyColors.Purple, OxyColors.Brown, OxyColors.DarkCyan };
                int ci = 0;

                // 全局均值参考线
                double globalMean = data.Values.SelectMany(pts => pts).Average(p => p.Mean);
                var meanLine = new LineSeries { Color = OxyColors.Gray, LineStyle = LineStyle.Dash, StrokeThickness = 1 };
                // 只要横跨所有因子的范围即可
                double xMin = double.MaxValue, xMax = double.MinValue;

                foreach (var kv in data)
                {
                    var series = new LineSeries
                    {
                        Title = kv.Key,
                        Color = colors[ci % colors.Length],
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 5,
                        StrokeThickness = 2
                    };
                    for (int i = 0; i < kv.Value.Count; i++)
                    {
                        var pt = kv.Value[i];
                        double xVal;
                        if (pt.Level is double d) xVal = d;
                        else if (double.TryParse(pt.Level?.ToString(), out var parsed)) xVal = parsed;
                        else xVal = i; // 类别因子用索引

                        series.Points.Add(new DataPoint(xVal, pt.Mean));
                        xMin = Math.Min(xMin, xVal);
                        xMax = Math.Max(xMax, xVal);
                    }
                    model.Series.Add(series);
                    ci++;
                }

                // 全局均值线
                if (xMin < xMax)
                {
                    meanLine.Points.Add(new DataPoint(xMin, globalMean));
                    meanLine.Points.Add(new DataPoint(xMax, globalMean));
                    model.Series.Add(meanLine);
                }

                model.Legends.Add(new Legend
                {
                    LegendPosition = LegendPosition.RightTop,
                    LegendFontSize = 10,
                    IsLegendVisible = true
                });

                OlsMainEffectsPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "主效应图加载失败"); }
        }

        // ═══ ★ v8: 交互效应图 ═══

        private async Task LoadOlsInteractionEffectsAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await _analysisService.GetInteractionEffectsForResponseAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName);
                var data = JsonConvert.DeserializeObject<List<InteractionEffectData>>(json);
                if (data == null || data.Count == 0) return;

                // 取第一对因子的交互效应图（如果多对，只显示最有代表性的）
                var first = data[0];

                var model = new PlotModel
                {
                    Title = $"交互效应: {first.Factor1} × {first.Factor2}",
                    TitleFontSize = 12,
                    PlotMargins = new OxyThickness(50, 25, 15, 35)
                };
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = _selectedResponseName, FontSize = 10 });
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = first.Factor1, FontSize = 10 });

                // 按 f2 水平分组画线
                var grouped = first.Data.GroupBy(d => d.F2?.ToString() ?? "").ToList();
                var colors = new[] { OxyColors.Blue, OxyColors.Red, OxyColors.Green, OxyColors.Orange };
                int ci = 0;

                foreach (var group in grouped)
                {
                    var series = new LineSeries
                    {
                        Title = $"{first.Factor2}={group.Key}",
                        Color = colors[ci % colors.Length],
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 4,
                        StrokeThickness = 2
                    };
                    foreach (var pt in group.OrderBy(p =>
                    {
                        if (p.F1 is double d) return d;
                        if (double.TryParse(p.F1?.ToString(), out var parsed)) return parsed;
                        return 0.0;
                    }))
                    {
                        double xVal;
                        if (pt.F1 is double d) xVal = d;
                        else if (double.TryParse(pt.F1?.ToString(), out var parsed)) xVal = parsed;
                        else xVal = 0;
                        series.Points.Add(new DataPoint(xVal, pt.Mean));
                    }
                    model.Series.Add(series);
                    ci++;
                }

                model.Legends.Add(new Legend
                {
                    LegendPosition = LegendPosition.RightTop,
                    LegendFontSize = 10,
                    IsLegendVisible = true
                });

                OlsInteractionPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "交互效应图加载失败"); }
        }

        // ═══ ★ v8: Tukey HSD ═══

        private async Task LoadTukeyHSDAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;

            // 检查是否有类别因子
            var batch = await _repository.GetBatchWithDetailsAsync(SelectedOlsBatch.BatchId);
            if (batch == null || !batch.Factors.Any(f => f.IsCategorical))
            {
                HasTukeyResult = false;
                return;
            }

            try
            {
                var catFactor = batch.Factors.First(f => f.IsCategorical).FactorName;
                var json = await _analysisService.GetTukeyHSDAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName, catFactor);
                var data = JsonConvert.DeserializeObject<TukeyHSDResult>(json);
                if (data == null || data.Error != null)
                {
                    HasTukeyResult = false;
                    return;
                }

                TukeyFactorName = data.FactorName;
                TukeyComparisons = data.Comparisons.Select(c => new TukeyComparisonItem
                {
                    Group1 = c.Group1,
                    Group2 = c.Group2,
                    MeanDiff = c.MeanDiff,
                    PValue = c.PValue,
                    CILower = c.CILower,
                    CIUpper = c.CIUpper,
                    Significant = c.Significant
                }).ToList();

                _tukeyGroupMeans = data.GroupMeans;
                HasTukeyResult = true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Tukey HSD 加载失败");
                HasTukeyResult = false;
            }
        }

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
        private async Task LoadBoxCoxAnalysisAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await _analysisService.GetBoxCoxAnalysisAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName);
                var data = JsonConvert.DeserializeObject<BoxCoxResult>(json);
                if (data == null || data.Error != null)
                {
                    HasBoxCoxResult = false;
                    BoxCoxRecommendation = data?.Error ?? "";
                    return;
                }

                BoxCoxLambda = data.RoundedLambda;
                BoxCoxTransformName = data.TransformName;
                BoxCoxOrigR2 = data.OriginalRSquared;
                BoxCoxTransR2 = data.TransformedRSquared;
                BoxCoxRecommend = data.RecommendTransform;
                BoxCoxRecommendation = data.Recommendation;
                HasBoxCoxResult = true;

                // λ profile 图
                if (data.LambdaProfile?.Lambdas != null)
                {
                    var model = new PlotModel
                    {
                        Title = "Box-Cox λ 优化曲线",
                        TitleFontSize = 11,
                        PlotMargins = new OxyThickness(50, 25, 15, 35)
                    };
                    model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "λ", FontSize = 10 });
                    model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "Log-Likelihood", FontSize = 10 });

                    var line = new LineSeries { Color = OxyColors.SteelBlue, StrokeThickness = 2 };
                    for (int i = 0; i < data.LambdaProfile.Lambdas.Count; i++)
                    {
                        var ll = data.LambdaProfile.LogLikelihoods[i];
                        if (ll.HasValue)
                            line.Points.Add(new DataPoint(data.LambdaProfile.Lambdas[i], ll.Value));
                    }
                    model.Series.Add(line);

                    // 最优 λ 标记
                    model.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                    {
                        Type = OxyPlot.Annotations.LineAnnotationType.Vertical,
                        X = data.OptimalLambda,
                        Color = OxyColors.Red,
                        LineStyle = LineStyle.Dash,
                        StrokeThickness = 1.5,
                        Text = $"λ={data.OptimalLambda:F2}",
                        TextColor = OxyColors.Red,
                        FontSize = 10
                    });

                    // CI 范围
                    if (data.LambdaCI != null && data.LambdaCI.Count == 2)
                    {
                        model.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                        {
                            Type = OxyPlot.Annotations.LineAnnotationType.Vertical,
                            X = data.LambdaCI[0],
                            Color = OxyColor.FromArgb(100, 0, 100, 255),
                            LineStyle = LineStyle.Dot,
                            StrokeThickness = 1
                        });
                        model.Annotations.Add(new OxyPlot.Annotations.LineAnnotation
                        {
                            Type = OxyPlot.Annotations.LineAnnotationType.Vertical,
                            X = data.LambdaCI[1],
                            Color = OxyColor.FromArgb(100, 0, 100, 255),
                            LineStyle = LineStyle.Dot,
                            StrokeThickness = 1
                        });
                    }

                    BoxCoxProfilePlot = model;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Box-Cox 分析失败");
                HasBoxCoxResult = false;
            }
        }

        private async Task ApplyBoxCoxAsync()
        {
            if (SelectedOlsBatch == null || !HasBoxCoxResult) return;
            try
            {
                IsLoading = true;
                OlsStatusText = $"正在应用 Box-Cox 变换 (λ={BoxCoxLambda})...";

                OlsResult = await _analysisService.ApplyBoxCoxAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName, BoxCoxLambda);

                if (OlsResult?.ModelSummary != null)
                {
                    OlsStatusText = $"Box-Cox 变换后: R²={OlsResult.ModelSummary.RSquared:F4}, {BoxCoxTransformName}";
                    UpdateEquationsDisplay(OlsResult);
                    await LoadOlsParetoAsync();
                    await LoadResidualDiagnosticsAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "应用 Box-Cox 失败");
                OlsStatusText = $"Box-Cox 应用失败: {ex.Message}";
            }
            finally { IsLoading = false; }
        }

        private async Task ExportOlsReportAsync()
        {
            if (SelectedOlsBatch == null || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var defaultName = $"OLS报告_{_selectedResponseName}_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                var path = _dialogService.ShowSaveFileDialog("Word 文件 (*.docx)|*.docx", ".docx", defaultName);
                if (string.IsNullOrEmpty(path)) return;

                IsLoading = true;
                OlsStatusText = "正在生成报告...";

                var resultJson = await _analysisService.ExportOlsReportAsync(
                    SelectedOlsBatch.BatchId, _selectedResponseName, path,
                    $"{_selectedResponseName} OLS 回归分析报告");

                var result = JsonConvert.DeserializeObject<Dictionary<string, object>>(resultJson);
                if (result?.ContainsKey("success") == true && (bool)result["success"])
                {
                    OlsStatusText = $"报告已导出: {path}";
                    _dialogService.ShowInfo($"OLS 分析报告已保存至:\n{path}", "报告已生成");
                }
                else
                {
                    var error = result?.ContainsKey("error") == true ? result["error"]?.ToString() : "未知错误";
                    OlsStatusText = $"报告导出失败: {error}";
                    _dialogService.ShowError($"报告导出失败: {error}", "错误");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "导出报告失败");
                _dialogService.ShowError($"导出失败: {ex.Message}", "错误");
            }
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
                    OlsStatusText = $"OLS 分析完成: R²=...";

                    // ★ v7: 不可估计项警告
                    InestimableWarning = OlsResult.InestimableWarning ?? "";
                    HasInestimableWarning = !string.IsNullOrEmpty(InestimableWarning);

                    // ★ v7: 初始化 OLS 曲面因子选择列表（仅连续因子可做曲面）
                    var batch = await _repository.GetBatchWithDetailsAsync(SelectedOlsBatch!.BatchId);
                    if (batch != null)
                    {
                        OlsContinuousFactors = batch.Factors
                            .Where(f => !f.IsCategorical)
                            .Select(f => f.FactorName)
                            .ToList();
                        if (OlsContinuousFactors.Count >= 2)
                        {
                            OlsSurfaceFactor1 = OlsContinuousFactors[0];
                            OlsSurfaceFactor2 = OlsContinuousFactors[1];
                        }
                    }

                    // 方程展示
                    UpdateEquationsDisplay(OlsResult);
                    // Pareto 图
                    await LoadOlsParetoAsync();
                    // 残差四合一
                    await LoadResidualDiagnosticsAsync();
                    // ★ v7: 自动加载 OLS 响应曲面
                    await LoadOlsSurfaceAsync();
                    // ★ v8: 加载主效应图 + 交互效应图 + 异常点分析 + Tukey HSD
                    await LoadOlsMainEffectsAsync();
                    await LoadOlsInteractionEffectsAsync();
                    await LoadOutlierAnalysisAsync();
                    await LoadTukeyHSDAsync();

                    await LoadBoxCoxAnalysisAsync();
                }
                else { OlsStatusText = "OLS 分析返回空结果"; }
            }
            catch (Exception ex) { _logger.LogError(ex, "OLS 分析失败"); OlsStatusText = $"OLS 分析失败: {ex.Message}"; }
            finally { IsLoading = false; }
        }
        // ═══ 新增方法: 解析方程展示 ═══

        private void UpdateEquationsDisplay(OLSAnalysisResult result)
        {
            EquationsInfo? eqInfo;
            if (IsCodedEquation)
            {
                eqInfo = result.ModelSummary?.Equations;
            }
            else
            {
                eqInfo = result.UncodedEquations;
            }

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

            // ★ v13: 纯连续因子时也要切换方程文本
            if (IsCodedEquation)
            {
                CurrentEquationText = result.ModelSummary?.Equation ?? "";
            }
            else
            {
                CurrentEquationText = result.UncodedEquation ?? result.ModelSummary?.Equation ?? "";
            }
        }
        // ═══ Pareto 图构建（支持点击切换） ═══

        private async Task LoadOlsParetoAsync()
        {
            try
            {
                var json = await _analysisService.GetEffectsParetoAsync(
                    SelectedOlsBatch!.BatchId, _selectedResponseName);
                var wrapper = JsonConvert.DeserializeObject<ParetoResult>(json);
                var data = wrapper?.Effects;
                _paretoLogWorthCrit = wrapper?.LogWorthCrit ?? 1.301;
                _paretoAlpha = wrapper?.Alpha ?? 0.05;
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
        /// <summary>
        /// 根据 _paretoTerms 的 IsIncluded 状态重新构建 OxyPlot 图
        /// 优化版: 自适应左边距、标签外置、显著性双色、精确 t 临界值
        /// </summary>
        private void RebuildParetoPlot()
        {
            var sorted = _paretoTerms.OrderByDescending(t => t.AbsT).ToList();
            int termCount = sorted.Count;

            int maxLabelLen = sorted.Max(t => t.TermName.Length);
            double leftMargin = Math.Max(100, Math.Min(220, maxLabelLen * 7.5 + 20));

            var model = new PlotModel
            {
                Title = $"效应显著性 Pareto 图 (α = {_paretoAlpha})",
                TitleFontSize = 12,
                TitleFontWeight = OxyPlot.FontWeights.Bold,
                Subtitle = "点击条形可切换保留/删除 — 蓝色=保留, 灰色=删除",
                SubtitleFontSize = 9,
                SubtitleColor = OxyColor.FromRgb(140, 140, 140),
                PlotMargins = new OxyThickness(leftMargin, 35, 60, 50),
                PlotAreaBorderThickness = new OxyThickness(1, 0, 0, 1),
                PlotAreaBorderColor = OxyColor.FromRgb(200, 200, 200)
            };

            var catAxis = new CategoryAxis
            {
                Position = AxisPosition.Left,
                GapWidth = termCount > 12 ? 0.15 : 0.3,
                FontSize = termCount > 15 ? 9.5 : 10.5,
                TextColor = OxyColor.FromRgb(60, 60, 60),
                TickStyle = TickStyle.None,
                AxislineStyle = LineStyle.None
            };
            for (int i = sorted.Count - 1; i >= 0; i--)
                catAxis.Labels.Add(sorted[i].TermName);

            double tCrit = _paretoLogWorthCrit;
            double maxT = sorted.Count > 0 ? sorted[0].AbsT : 5;
            double xMax = Math.Max(maxT, tCrit) * 1.25;

            var valAxis = new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = "LogWorth = -log\u2081\u2080(P)",
                TitleFontSize = 10,
                Minimum = 0,
                AbsoluteMinimum = 0,
                Maximum = xMax,
                FontSize = 10,
                MajorGridlineStyle = LineStyle.Dot,
                MajorGridlineColor = OxyColor.FromRgb(230, 230, 230),
                TickStyle = TickStyle.Outside
            };

            model.Axes.Add(catAxis);
            model.Axes.Add(valAxis);

            var series = new BarSeries
            {
                StrokeThickness = 0.8,
                StrokeColor = OxyColor.FromRgb(160, 160, 160),
                LabelPlacement = LabelPlacement.Outside,
                LabelFormatString = "{0:F2}",
                LabelMargin = 4,
                FontSize = termCount > 15 ? 8 : 9
            };

            for (int i = sorted.Count - 1; i >= 0; i--)
            {
                var item = sorted[i];
                OxyColor color;
                if (!item.IsIncluded)
                    color = OxyColor.FromArgb(100, 200, 200, 200);
                else if (item.IsSignificant)
                    color = OxyColor.FromRgb(66, 133, 244);
                else
                    color = OxyColor.FromRgb(160, 196, 255);
                series.Items.Add(new BarItem { Value = item.AbsT, Color = color });
            }
            model.Series.Add(series);

            // ── 红色虚线 + 文字（放在线的中间位置，避开条形和边界裁切）──
            var critLine = new LineAnnotation
            {
                Type = LineAnnotationType.Vertical,
                X = tCrit,
                Color = OxyColor.FromArgb(200, 220, 50, 50),
                LineStyle = LineStyle.Dash,
                StrokeThickness = 1.8,
                Text = $"  {_paretoLogWorthCrit:F2}",
                TextColor = OxyColor.FromRgb(200, 50, 50),
                FontSize = 9,
                FontWeight = OxyPlot.FontWeights.Bold,
                TextLinePosition = 0.5,
                TextOrientation = AnnotationTextOrientation.Vertical,
                TextHorizontalAlignment = HorizontalAlignment.Left,
                TextMargin = 3
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

                    if (OlsResult.Coefficients != null)
                    {
                        foreach (var term in _paretoTerms.Where(t => t.IsIncluded))
                        {
                            var newCoeff = OlsResult.Coefficients
                                .FirstOrDefault(c => c.Term == term.TermName);
                            if (newCoeff != null)
                            {
                                term.AbsT = Math.Abs(newCoeff.TValue);
                                term.PValue = newCoeff.PValue;
                                term.IsSignificant = newCoeff.PValue < 0.05;
                            }
                        }
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
            if (string.IsNullOrEmpty(_selectedResponseName)) return;

            // ★ 直接导入模式和批次模式都要支持
            if (!IsDirectImportMode && SelectedOlsBatch == null) return;

            try
            {
                IsLoading = true;
                OlsStatusText = "正在生成预测刻画器...";

                string json;
                if (IsDirectImportMode)
                {
                    json = await Task.Run(() =>
                        _analysisService.GetPredictionProfilerDirectAsync(_selectedResponseName));
                }
                else
                {
                    json = await _analysisService.GetPredictionProfilerAsync(
                        SelectedOlsBatch!.BatchId, _selectedResponseName);
                }

                var data = JsonConvert.DeserializeObject<ProfilerResult>(json);
                if (data?.Factors == null) return;

                _lastProfilerData = data;
                _profilerFactorOrder = data.Factors.Keys.ToList();

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
            if (!IsProfilerLoaded) return;
            if (!IsDirectImportMode && SelectedOlsBatch == null) return;

            _profilerCurrentValues[factorName] = newValue;

            try
            {
                var fixedValues = new Dictionary<string, object>(_profilerCurrentValues);
                string json;

                if (IsDirectImportMode)
                {
                    var fixedJson = JsonConvert.SerializeObject(fixedValues);
                    json = await Task.Run(() =>
                        _analysisService.GetPredictionProfilerDirectAsync(
                            _selectedResponseName, 50, fixedJson));
                }
                else
                {
                    json = await _analysisService.GetPredictionProfilerAsync(
                        SelectedOlsBatch!.BatchId, _selectedResponseName, 50, fixedValues);
                }

                var data = JsonConvert.DeserializeObject<ProfilerResult>(json);
                if (data?.Factors == null) return;

                _lastProfilerData = data;
                foreach (var kv in data.Factors)
                    _profilerCurrentValues[kv.Key] = kv.Value.CurrentValue ?? _profilerCurrentValues[kv.Key];

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

                // ★ v7: 类别因子 CI 带
                if (fdata.YLower != null && fdata.YUpper != null
                    && fdata.YLower.Count == fdata.Y.Count && fdata.Y.Count > 0)
                {
                    var areaSeries = new AreaSeries
                    {
                        Color = OxyColor.FromArgb(0, 0, 0, 0),
                        Fill = OxyColor.FromArgb(40, 70, 130, 220),
                        StrokeThickness = 0
                    };
                    for (int j = 0; j < fdata.Y.Count && j < labels.Count; j++)
                    {
                        areaSeries.Points.Add(new DataPoint(j, fdata.YUpper[j]));
                        areaSeries.Points2.Add(new DataPoint(j, fdata.YLower[j]));
                    }
                    pm.Series.Add(areaSeries);
                }

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

                // ★ v7: 连续因子 CI 带
                if (fdata.YLower != null && fdata.YUpper != null
                    && fdata.YLower.Count == fdata.Y.Count && fdata.Y.Count > 0)
                {
                    var areaSeries = new AreaSeries
                    {
                        Color = OxyColor.FromArgb(0, 0, 0, 0),
                        Fill = OxyColor.FromArgb(40, 70, 130, 220),
                        StrokeThickness = 0
                    };
                    for (int j = 0; j < fdata.X.Count; j++)
                    {
                        areaSeries.Points.Add(new DataPoint(fdata.X[j], fdata.YUpper[j]));
                        areaSeries.Points2.Add(new DataPoint(fdata.X[j], fdata.YLower[j]));
                    }
                    pm.Series.Add(areaSeries);
                }

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

                pm.Axes.Add(new LinearAxis
                {
                    Position = AxisPosition.Left,
                    FontSize = 9,
                    Minimum = yMin - yPad,
                    Maximum = yMax + yPad,
                    Title = plots.Count == 0 ? data.ResponseName : ""
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

                    // ★ v7: 类别因子置信区间带
                    if (fdata.YLower != null && fdata.YUpper != null
                        && fdata.YLower.Count == fdata.Y.Count && fdata.Y.Count > 0)
                    {
                        var areaSeries = new AreaSeries
                        {
                            Color = OxyColor.FromArgb(0, 0, 0, 0),
                            Fill = OxyColor.FromArgb(40, 70, 130, 220),
                            StrokeThickness = 0
                        };
                        for (int j = 0; j < fdata.Y.Count && j < labels.Count; j++)
                        {
                            areaSeries.Points.Add(new DataPoint(j, fdata.YUpper[j]));
                            areaSeries.Points2.Add(new DataPoint(j, fdata.YLower[j]));
                        }
                        pm.Series.Add(areaSeries);
                    }

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
                        var highlight = new ScatterSeries
                        {
                            MarkerType = MarkerType.Circle,
                            MarkerSize = 7,
                            MarkerFill = OxyColors.Red
                        };
                        highlight.Points.Add(new ScatterPoint(curIdx, fdata.Y[curIdx]));
                        pm.Series.Add(highlight);
                    }

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
                    // 连续因子
                    pm.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, FontSize = 9 });

                    // ★ v7: 连续因子置信区间带（必须在曲线之前添加）
                    if (fdata.YLower != null && fdata.YUpper != null
                        && fdata.YLower.Count == fdata.Y.Count && fdata.Y.Count > 0)
                    {
                        var areaSeries = new AreaSeries
                        {
                            Color = OxyColor.FromArgb(0, 0, 0, 0),
                            Fill = OxyColor.FromArgb(40, 70, 130, 220),
                            StrokeThickness = 0
                        };
                        for (int j = 0; j < fdata.X.Count; j++)
                        {
                            areaSeries.Points.Add(new DataPoint(fdata.X[j], fdata.YUpper[j]));
                            areaSeries.Points2.Add(new DataPoint(fdata.X[j], fdata.YLower[j]));
                        }
                        pm.Series.Add(areaSeries);
                    }

                    var line = new LineSeries { Color = OxyColors.SteelBlue, StrokeThickness = 2 };
                    for (int i = 0; i < fdata.X.Count; i++)
                        line.Points.Add(new DataPoint(fdata.X[i], fdata.Y[i]));
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
            if (string.IsNullOrEmpty(_selectedResponseName)) return;
            if (!IsDirectImportMode && SelectedOlsBatch == null) return;

            try
            {
                IsLoading = true;
                OlsStatusText = maximize ? "正在搜索最大值..." : "正在搜索最小值...";

                string json;
                if (IsDirectImportMode)
                {
                    json = await Task.Run(() =>
                        _analysisService.FindOptimalDirectAsync(_selectedResponseName, maximize));
                }
                else
                {
                    json = await _analysisService.FindOptimalAsync(
                        SelectedOlsBatch!.BatchId, _selectedResponseName, maximize);
                }

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
                        foreach (var kv in data.OptimalFactors)
                            _profilerCurrentValues[kv.Key] = kv.Value;

                        try
                        {
                            string profilerJson;
                            if (IsDirectImportMode)
                            {
                                var fixedJson = JsonConvert.SerializeObject(data.OptimalFactors);
                                profilerJson = await Task.Run(() =>
                                    _analysisService.GetPredictionProfilerDirectAsync(
                                        _selectedResponseName, 50, fixedJson));
                            }
                            else
                            {
                                profilerJson = await _analysisService.GetPredictionProfilerAsync(
                                    SelectedOlsBatch!.BatchId, _selectedResponseName, 50, data.OptimalFactors);
                            }

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

        // ══════════════ OLS 直接导入 Excel 数据分析 ══════════════

        /// <summary>
        /// ★ 新增: 在 OLS Tab 直接导入 Excel 做分析，不依赖任何 DOE 批次
        /// Excel 格式: 第一行列头 = 因子名 + 响应名，后续行为数据
        /// 自动检测: 最后一列或多列为纯数值 → 响应变量；其他列 → 因子
        /// </summary>
        private async Task OlsImportExcelAsync()
        {
            try
            {
                var path = _dialogService.ShowOpenFileDialog("Excel 文件 (*.xlsx)|*.xlsx");
                if (string.IsNullOrEmpty(path)) return;

                IsLoading = true;
                OlsStatusText = "正在读取 Excel 文件...";

                // Step 1: 后台读取原始数据
                var rawData = await Task.Run(() => ReadExcelRawData(path));

                if (rawData.Errors.Count > 0)
                {
                    _dialogService.ShowError($"读取失败:\n{string.Join("\n", rawData.Errors)}", "导入失败");
                    OlsStatusText = "导入失败";
                    return;
                }

                if (rawData.DataRowCount < 3)
                {
                    _dialogService.ShowError($"有效数据不足 3 行（当前 {rawData.DataRowCount} 行）", "数据不足");
                    OlsStatusText = "数据不足";
                    return;
                }

                IsLoading = false;

                // Step 2: 构建预填配置，弹窗让用户确认
                var columnConfigs = BuildColumnConfigs(rawData);

                bool confirmed = false;
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    var dialog = new Views.OlsImportConfirmDialog(columnConfigs, path, rawData.DataRowCount);
                    dialog.Owner = System.Windows.Application.Current.MainWindow;
                    dialog.ShowDialog();
                    confirmed = dialog.Confirmed;
                });

                if (!confirmed)
                {
                    OlsStatusText = "已取消导入";
                    return;
                }

                // Step 3: 根据确认的配置构建分析数据
                IsLoading = true;
                OlsStatusText = "正在构建分析数据...";

                var importResult = ApplyConfirmedConfig(rawData, columnConfigs);

                if (importResult.DataCount < 3)
                {
                    _dialogService.ShowError("排除空值后有效数据不足 3 行。", "数据不足");
                    OlsStatusText = "数据不足";
                    return;
                }

                // Step 4: 保存并执行分析
                DirectImportFilePath = path;
                _directImportFactorNames = importResult.FactorNames;
                _directImportResponseNames = importResult.ResponseNames;
                _directImportFactorTypes = importResult.FactorTypes;
                _directImportFactorsData = importResult.FactorsData;
                _directImportResponsesData = importResult.ResponsesData;

                IsDirectImportMode = true;
                SelectedOlsBatch = null;

                ResponseNames = importResult.ResponseNames;
                _selectedResponseName = ResponseNames.FirstOrDefault() ?? "";
                RaisePropertyChanged(nameof(SelectedResponseName));

                OlsStatusText = $"已导入 {importResult.DataCount} 行，{importResult.FactorNames.Count} 个因子，{importResult.ResponseNames.Count} 个响应";

                await RunDirectOlsAnalysisAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OLS 直接导入失败");
                _dialogService.ShowError($"导入失败: {ex.Message}", "错误");
                OlsStatusText = $"导入失败: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }
        /// <summary>
        /// 根据原始数据自动预填列配置
        /// 规则: 含文字列→类别因子, 全数值列→连续因子, 最后一个全数值列→响应
        /// </summary>
        private static ObservableCollection<ColumnConfigItem> BuildColumnConfigs(DirectImportRawData rawData)
        {
            var configs = new ObservableCollection<ColumnConfigItem>();

            // 找最后一个全数值列作为默认响应
            int lastNumericCol = -1;
            for (int c = rawData.Headers.Count - 1; c >= 0; c--)
            {
                if (rawData.IsNumericColumn[c]) { lastNumericCol = c; break; }
            }

            for (int c = 0; c < rawData.Headers.Count; c++)
            {
                var data = rawData.ColumnData[c];
                var nonNullValues = data.Where(v => v != null).ToList();
                var uniqueCount = nonNullValues.Select(v => v!.ToString()).Distinct().Count();
                var preview = nonNullValues.Take(5).Select(v => v!.ToString() ?? "").ToList();

                var item = new ColumnConfigItem
                {
                    ColumnName = rawData.Headers[c],
                    IsAllNumeric = rawData.IsNumericColumn[c],
                    UniqueCount = uniqueCount,
                    PreviewValues = preview!
                };

                // 自动预填角色和类型
                if (c == lastNumericCol)
                {
                    item.Role = "Response";
                    item.FactorType = "";
                }
                else if (!rawData.IsNumericColumn[c])
                {
                    item.Role = "Factor";
                    item.FactorType = "Categorical";
                }
                else if (rawData.IsNumericColumn[c] && uniqueCount <= 5 && nonNullValues.Count > 10)
                {
                    // 数值列但唯一值很少 → 可能是类别，但默认连续让用户决定
                    item.Role = "Factor";
                    item.FactorType = "Continuous";
                }
                else
                {
                    item.Role = "Factor";
                    item.FactorType = "Continuous";
                }

                configs.Add(item);
            }

            return configs;
        }
        /// <summary>
        /// 根据用户确认的列配置，从原始数据构建 OLS 分析所需的结构化数据
        /// </summary>
        private static DirectImportResult ApplyConfirmedConfig(
            DirectImportRawData rawData,
            ObservableCollection<ColumnConfigItem> configs)
        {
            var result = new DirectImportResult();

            var factorConfigs = configs.Where(c => c.Role == "Factor").ToList();
            var responseConfigs = configs.Where(c => c.Role == "Response").ToList();

            result.FactorNames = factorConfigs.Select(c => c.ColumnName).ToList();
            result.ResponseNames = responseConfigs.Select(c => c.ColumnName).ToList();

            // 因子类型
            foreach (var fc in factorConfigs)
                result.FactorTypes[fc.ColumnName] = fc.FactorType == "Categorical" ? "categorical" : "continuous";

            // 列名→索引映射
            var colIndexMap = new Dictionary<string, int>();
            for (int i = 0; i < rawData.Headers.Count; i++)
                colIndexMap[rawData.Headers[i]] = i;

            // 构建数据行
            for (int row = 0; row < rawData.DataRowCount; row++)
            {
                // 检查所有必需列是否有值
                bool hasNull = false;
                foreach (var fc in factorConfigs)
                {
                    if (rawData.ColumnData[colIndexMap[fc.ColumnName]][row] == null) { hasNull = true; break; }
                }
                if (hasNull) continue;
                foreach (var rc in responseConfigs)
                {
                    if (rawData.ColumnData[colIndexMap[rc.ColumnName]][row] == null) { hasNull = true; break; }
                }
                if (hasNull) continue;

                // 因子行
                var factorRow = new Dictionary<string, object>();
                foreach (var fc in factorConfigs)
                {
                    factorRow[fc.ColumnName] = rawData.ColumnData[colIndexMap[fc.ColumnName]][row]!;
                }
                result.FactorsData.Add(factorRow);

                // 响应值
                foreach (var rc in responseConfigs)
                {
                    var respName = rc.ColumnName;
                    if (!result.ResponsesData.ContainsKey(respName))
                        result.ResponsesData[respName] = new List<double>();
                    result.ResponsesData[respName].Add((double)rawData.ColumnData[colIndexMap[respName]][row]!);
                }
            }

            result.DataCount = result.FactorsData.Count;
            return result;
        }
        /// <summary>
        /// 执行直接导入数据的 OLS 分析
        /// </summary>
        private async Task RunDirectOlsAnalysisAsync()
        {
            if (!IsDirectImportMode || _directImportFactorsData.Count == 0) return;
            if (string.IsNullOrEmpty(_selectedResponseName)) return;

            if (!_directImportResponsesData.TryGetValue(_selectedResponseName, out var responsesData))
            {
                OlsStatusText = $"未找到响应变量: {_selectedResponseName}";
                return;
            }

            try
            {
                IsLoading = true;
                OlsStatusText = "正在进行 OLS 回归分析...";

                // ═══ Step 1: FitOLS ═══
                OlsResult = await _analysisService.FitOlsDirectAsync(
                    _directImportFactorsData,
                    responsesData,
                    _selectedResponseName,
                    _directImportFactorTypes,
                    "quadratic");

                if (OlsResult?.ModelSummary == null)
                {
                    OlsStatusText = "OLS 分析返回空结果";
                    return;
                }

                OlsStatusText = $"OLS 分析完成: R²={OlsResult.ModelSummary.RSquared:F4}, " +
                                 $"R²adj={OlsResult.ModelSummary.RSquaredAdj:F4}";

                // 不可估计项警告
                InestimableWarning = OlsResult.InestimableWarning ?? "";
                HasInestimableWarning = !string.IsNullOrEmpty(InestimableWarning);

                // 连续因子列表（曲面图用）
                OlsContinuousFactors = _directImportFactorNames
                    .Where(f => _directImportFactorTypes.GetValueOrDefault(f, "continuous") == "continuous")
                    .ToList();
                if (OlsContinuousFactors.Count >= 2)
                {
                    OlsSurfaceFactor1 = OlsContinuousFactors[0];
                    OlsSurfaceFactor2 = OlsContinuousFactors[1];
                }

                // ═══ Step 2: 方程展示 ═══
                UpdateEquationsDisplay(OlsResult);

                // ═══ Step 3: Pareto 图 ═══
                await LoadDirectOlsParetoAsync();

                // ═══ Step 4: 残差四合一 ═══
                await LoadDirectResidualDiagnosticsAsync();

                // ═══ Step 5: 响应曲面 + 等高线 ═══
                await LoadDirectOlsSurfaceAsync();

                // ═══ Step 6: 主效应图 ═══
                await LoadDirectMainEffectsAsync();

                // ═══ Step 7: 交互效应图 ═══
                await LoadDirectInteractionEffectsAsync();

                // ═══ Step 8: 异常点分析 ═══
                await LoadDirectOutlierAnalysisAsync();

                // ═══ Step 9: Box-Cox 分析 ═══
                await LoadDirectBoxCoxAnalysisAsync();

                // ═══ Step 10: Tukey HSD（仅在有类别因子时） ═══
                if (_directImportFactorTypes.Values.Any(v => v == "categorical"))
                {
                    await LoadDirectTukeyHSDAsync();
                }
                else
                {
                    HasTukeyResult = false;
                }

                OlsStatusText = $"OLS 分析完成: R²={OlsResult.ModelSummary.RSquared:F4}, " +
                                 $"R²adj={OlsResult.ModelSummary.RSquaredAdj:F4}, " +
                                 $"来源=直接导入";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "直接导入 OLS 分析失败");
                OlsStatusText = $"分析失败: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        /// 直接导入模式下的 Pareto 图
        /// </summary>
        private async Task LoadDirectOlsParetoAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await Task.Run(() =>
                    _analysisService.GetEffectsParetoDirectAsync(_selectedResponseName));

                var wrapper = JsonConvert.DeserializeObject<ParetoResult>(json);
                var data = wrapper?.Effects;
                _paretoLogWorthCrit = wrapper?.LogWorthCrit ?? 1.301;
                _paretoAlpha = wrapper?.Alpha ?? 0.05;
                if (data == null || data.Count == 0) return;

                _paretoTerms = data.Select(d => new ParetoTermItem
                {
                    TermName = d.Term,
                    AbsT = d.AbsT,
                    PValue = d.PValue,
                    IsSignificant = d.Significant,
                    IsIncluded = true
                }).ToList();

                RebuildParetoPlot();
                UpdateTermsText();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "直接导入 Pareto 图加载失败");
            }
        }

        /// <summary>
        /// 直接导入模式下的残差诊断
        /// </summary>
        private async Task LoadDirectResidualDiagnosticsAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await Task.Run(() =>
                    _analysisService.GetResidualDiagnosticsDirectAsync(_selectedResponseName));

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
                    histSeries.Items.Add(new RectangleBarItem { X0 = x0, X1 = x1, Y0 = 0, Y1 = diag.ResidualsHistogram.Frequencies[i], Color = OxyColor.FromRgb(100, 149, 237) });
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
                if (diag.ResidualsVsOrder.Order.Count > 0)
                {
                    var zeroLine = new LineSeries { Color = OxyColors.Gray, LineStyle = LineStyle.Dash, StrokeThickness = 1 };
                    zeroLine.Points.Add(new DataPoint(1, 0));
                    zeroLine.Points.Add(new DataPoint(diag.ResidualsVsOrder.Order.Max(), 0));
                    rvoModel.Series.Add(zeroLine);
                }
                ResidVsOrderPlot = rvoModel;
            }
            catch (Exception ex) { _logger.LogError(ex, "直接导入残差诊断图加载失败"); }
        }
        // ═══ 直接导入模式: 响应曲面 ═══
        private async Task LoadDirectOlsSurfaceAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            if (string.IsNullOrEmpty(OlsSurfaceFactor1) || string.IsNullOrEmpty(OlsSurfaceFactor2)
                || OlsSurfaceFactor1 == OlsSurfaceFactor2) return;

            try
            {
                // 构建 bounds JSON
                var boundsDict = new Dictionary<string, object>();
                foreach (var factorRow in _directImportFactorsData)
                {
                    foreach (var kv in factorRow)
                    {
                        if (!boundsDict.ContainsKey(kv.Key))
                        {
                            if (_directImportFactorTypes.GetValueOrDefault(kv.Key, "continuous") == "categorical")
                            {
                                boundsDict[kv.Key] = _directImportFactorsData
                                    .Select(r => r[kv.Key]?.ToString() ?? "").Distinct().OrderBy(v => v).ToList();
                            }
                            else
                            {
                                var vals = _directImportFactorsData.Select(r => Convert.ToDouble(r[kv.Key])).ToList();
                                boundsDict[kv.Key] = new[] { vals.Min(), vals.Max() };
                            }
                        }
                    }
                }
                var boundsJson = JsonConvert.SerializeObject(boundsDict);

                // 3D 曲面图
                var imageBytes = await Task.Run(() =>
                    _analysisService.GetOlsSurfaceImageDirectAsync(OlsSurfaceFactor1, OlsSurfaceFactor2, boundsJson));

                if (imageBytes.Length > 0)
                {
                    var bmp = new BitmapImage();
                    using (var ms = new MemoryStream(imageBytes))
                    {
                        bmp.BeginInit(); bmp.CacheOption = BitmapCacheOption.OnLoad;
                        bmp.StreamSource = ms; bmp.EndInit(); bmp.Freeze();
                    }
                    System.Windows.Application.Current.Dispatcher.Invoke(() => { OlsSurfaceImage = bmp; });
                }
                else { OlsSurfaceImage = null; }

                // 等高线
                var contourJson = await Task.Run(() =>
                    _analysisService.GetOlsContourDataDirectAsync(OlsSurfaceFactor1, OlsSurfaceFactor2, boundsJson));
                var contourData = JsonConvert.DeserializeObject<SurfaceData>(contourJson);
                if (contourData?.Z != null) BuildContourPlot(contourData);
            }
            catch (Exception ex) { _logger.LogError(ex, "直接导入 OLS 响应曲面加载失败"); }
        }

        // ═══ 直接导入模式: 主效应图 ═══
        private async Task LoadDirectMainEffectsAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await Task.Run(() =>
                    _analysisService.GetMainEffectsDirectAsync(_selectedResponseName));
                var data = JsonConvert.DeserializeObject<Dictionary<string, List<MainEffectPoint>>>(json);
                if (data == null || data.Count == 0) return;

                var model = new PlotModel
                {
                    Title = "主效应图",
                    TitleFontSize = 12,
                    PlotMargins = new OxyThickness(50, 25, 15, 35)
                };
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = _selectedResponseName, FontSize = 10 });
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "因子水平", FontSize = 10 });

                var colors = new[] { OxyColors.Blue, OxyColors.Red, OxyColors.Green, OxyColors.Orange,
                             OxyColors.Purple, OxyColors.Brown, OxyColors.DarkCyan };
                int ci = 0;
                double globalMean = data.Values.SelectMany(pts => pts).Average(p => p.Mean);
                double xMin = double.MaxValue, xMax = double.MinValue;

                foreach (var kv in data)
                {
                    var series = new LineSeries
                    {
                        Title = kv.Key,
                        Color = colors[ci % colors.Length],
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 5,
                        StrokeThickness = 2
                    };
                    for (int i = 0; i < kv.Value.Count; i++)
                    {
                        var pt = kv.Value[i];
                        double xVal = pt.Level is double d ? d
                            : double.TryParse(pt.Level?.ToString(), out var parsed) ? parsed : i;
                        series.Points.Add(new DataPoint(xVal, pt.Mean));
                        xMin = Math.Min(xMin, xVal); xMax = Math.Max(xMax, xVal);
                    }
                    model.Series.Add(series); ci++;
                }

                if (xMin < xMax)
                {
                    var meanLine = new LineSeries { Color = OxyColors.Gray, LineStyle = LineStyle.Dash, StrokeThickness = 1 };
                    meanLine.Points.Add(new DataPoint(xMin, globalMean));
                    meanLine.Points.Add(new DataPoint(xMax, globalMean));
                    model.Series.Add(meanLine);
                }

                model.Legends.Add(new Legend { LegendPosition = LegendPosition.RightTop, LegendFontSize = 10, IsLegendVisible = true });
                OlsMainEffectsPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "直接导入主效应图加载失败"); }
        }

        // ═══ 直接导入模式: 交互效应图 ═══
        private async Task LoadDirectInteractionEffectsAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await Task.Run(() =>
                    _analysisService.GetInteractionEffectsDirectAsync(_selectedResponseName));
                var data = JsonConvert.DeserializeObject<List<InteractionEffectData>>(json);
                if (data == null || data.Count == 0) return;

                var first = data[0];
                var model = new PlotModel
                {
                    Title = $"交互效应: {first.Factor1} × {first.Factor2}",
                    TitleFontSize = 12,
                    PlotMargins = new OxyThickness(50, 25, 15, 35)
                };
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = _selectedResponseName, FontSize = 10 });
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = first.Factor1, FontSize = 10 });

                var grouped = first.Data.GroupBy(d => d.F2?.ToString() ?? "").ToList();
                var colors = new[] { OxyColors.Blue, OxyColors.Red, OxyColors.Green, OxyColors.Orange };
                int ci = 0;

                foreach (var group in grouped)
                {
                    var series = new LineSeries
                    {
                        Title = $"{first.Factor2}={group.Key}",
                        Color = colors[ci % colors.Length],
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 4,
                        StrokeThickness = 2
                    };
                    foreach (var pt in group.OrderBy(p => p.F1 is double d ? d : double.TryParse(p.F1?.ToString(), out var parsed) ? parsed : 0.0))
                    {
                        double xVal = pt.F1 is double d ? d : double.TryParse(pt.F1?.ToString(), out var parsed) ? parsed : 0;
                        series.Points.Add(new DataPoint(xVal, pt.Mean));
                    }
                    model.Series.Add(series); ci++;
                }

                model.Legends.Add(new Legend { LegendPosition = LegendPosition.RightTop, LegendFontSize = 10, IsLegendVisible = true });
                OlsInteractionPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "直接导入交互效应图加载失败"); }
        }

        // ═══ 直接导入模式: 异常点分析 ═══
        private async Task LoadDirectOutlierAnalysisAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await Task.Run(() =>
                    _analysisService.GetOutlierAnalysisDirectAsync(_selectedResponseName));
                var data = JsonConvert.DeserializeObject<OutlierAnalysisResult>(json);
                if (data == null) return;

                OutlierItems = data.Outliers.Select(o => new OutlierItem
                {
                    Index = o.Index,
                    Actual = o.Actual,
                    Predicted = o.Predicted,
                    Residual = o.Residual,
                    StdResidual = o.StdResidual,
                    CooksD = o.CooksD,
                    Leverage = o.Leverage,
                    Reasons = string.Join("; ", o.Reasons),
                    IsExcluded = false
                }).ToList();

                HasOutliers = data.OutlierCount > 0;
                OutlierSummaryText = data.OutlierCount > 0
                    ? $"检测到 {data.OutlierCount} 个异常点（共 {data.TotalObservations} 个观测）"
                    : $"未检测到异常点（共 {data.TotalObservations} 个观测）";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "直接导入异常点分析失败");
                HasOutliers = false;
            }
        }

        // ═══ 直接导入模式: Box-Cox ═══
        private async Task LoadDirectBoxCoxAnalysisAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            try
            {
                var json = await Task.Run(() =>
                    _analysisService.GetBoxCoxAnalysisDirectAsync(_selectedResponseName));
                var data = JsonConvert.DeserializeObject<BoxCoxResult>(json);
                if (data == null || data.Error != null)
                {
                    HasBoxCoxResult = false;
                    BoxCoxRecommendation = data?.Error ?? "";
                    return;
                }

                BoxCoxLambda = data.RoundedLambda;
                BoxCoxTransformName = data.TransformName;
                BoxCoxOrigR2 = data.OriginalRSquared;
                BoxCoxTransR2 = data.TransformedRSquared;
                BoxCoxRecommend = data.RecommendTransform;
                BoxCoxRecommendation = data.Recommendation;
                HasBoxCoxResult = true;

                if (data.LambdaProfile?.Lambdas != null)
                {
                    var model = new PlotModel
                    {
                        Title = "Box-Cox λ 优化曲线",
                        TitleFontSize = 11,
                        PlotMargins = new OxyThickness(50, 25, 15, 35)
                    };
                    model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "λ", FontSize = 10 });
                    model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "Log-Likelihood", FontSize = 10 });

                    var line = new LineSeries { Color = OxyColors.SteelBlue, StrokeThickness = 2 };
                    for (int i = 0; i < data.LambdaProfile.Lambdas.Count; i++)
                    {
                        if (data.LambdaProfile.LogLikelihoods[i].HasValue)
                            line.Points.Add(new DataPoint(data.LambdaProfile.Lambdas[i], data.LambdaProfile.LogLikelihoods[i]!.Value));
                    }
                    model.Series.Add(line);

                    model.Annotations.Add(new LineAnnotation
                    {
                        Type = LineAnnotationType.Vertical,
                        X = data.OptimalLambda,
                        Color = OxyColors.Red,
                        LineStyle = LineStyle.Dash,
                        StrokeThickness = 1.5,
                        Text = $"λ={data.OptimalLambda:F2}",
                        TextColor = OxyColors.Red,
                        FontSize = 10
                    });

                    if (data.LambdaCI != null && data.LambdaCI.Count == 2)
                    {
                        model.Annotations.Add(new LineAnnotation { Type = LineAnnotationType.Vertical, X = data.LambdaCI[0], Color = OxyColor.FromArgb(100, 0, 100, 255), LineStyle = LineStyle.Dot, StrokeThickness = 1 });
                        model.Annotations.Add(new LineAnnotation { Type = LineAnnotationType.Vertical, X = data.LambdaCI[1], Color = OxyColor.FromArgb(100, 0, 100, 255), LineStyle = LineStyle.Dot, StrokeThickness = 1 });
                    }

                    BoxCoxProfilePlot = model;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "直接导入 Box-Cox 分析失败");
                HasBoxCoxResult = false;
            }
        }

        // ═══ 直接导入模式: Tukey HSD ═══
        private async Task LoadDirectTukeyHSDAsync()
        {
            if (!IsDirectImportMode || string.IsNullOrEmpty(_selectedResponseName)) return;
            var catFactor = _directImportFactorTypes.FirstOrDefault(kv => kv.Value == "categorical").Key;
            if (string.IsNullOrEmpty(catFactor)) { HasTukeyResult = false; return; }

            try
            {
                var json = await Task.Run(() =>
                    _analysisService.GetTukeyHSDDirectAsync(_selectedResponseName, catFactor));

                var data = JsonConvert.DeserializeObject<TukeyHSDResult>(json);
                if (data == null || data.Error != null) { HasTukeyResult = false; return; }

                TukeyFactorName = data.FactorName;
                TukeyComparisons = data.Comparisons.Select(c => new TukeyComparisonItem
                {
                    Group1 = c.Group1,
                    Group2 = c.Group2,
                    MeanDiff = c.MeanDiff,
                    PValue = c.PValue,
                    CILower = c.CILower,
                    CIUpper = c.CIUpper,
                    Significant = c.Significant
                }).ToList();
                _tukeyGroupMeans = data.GroupMeans;
                HasTukeyResult = true;
            }
            catch (Exception ex) { _logger.LogError(ex, "直接导入 Tukey HSD 失败"); HasTukeyResult = false; }
        }
        /// <summary>
        /// 读取 Excel 文件用于直接 OLS 分析
        /// 自动识别: 尝试将每列转为数值，全部可转 → 连续因子/响应；否则 → 类别因子
        /// 规则: 最后 N 列全为数值 → 作为响应变量，其余为因子
        /// 但如果 Excel 只有 2 列数值，则最后一列为响应
        /// </summary>
        /// <summary>
        /// 读取 Excel 文件的原始列数据（不做角色判断，由弹窗确认）
        /// </summary>
        private static DirectImportRawData ReadExcelRawData(string path)
        {
            var result = new DirectImportRawData();
            try
            {
                using var package = new ExcelPackage(new FileInfo(path));
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null || ws.Dimension == null)
                {
                    result.Errors.Add("Excel 文件为空或没有工作表");
                    return result;
                }

                int totalCols = ws.Dimension.End.Column;
                int totalRows = ws.Dimension.End.Row;

                if (totalRows < 2)
                {
                    result.Errors.Add("至少需要 1 行列头 + 1 行数据");
                    return result;
                }

                // 读取列头
                for (int col = 1; col <= totalCols; col++)
                {
                    var header = ws.Cells[1, col].Text?.Trim();
                    if (string.IsNullOrEmpty(header)) header = $"Col{col}";
                    result.Headers.Add(header);
                    result.ColumnData.Add(new List<object?>());
                    result.IsNumericColumn.Add(true);
                }

                // 读取数据
                for (int row = 2; row <= totalRows; row++)
                {
                    bool allEmpty = true;
                    for (int c = 0; c < totalCols; c++)
                    {
                        var cellText = ws.Cells[row, c + 1].Text?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(cellText)) allEmpty = false;

                        if (double.TryParse(cellText, out var numVal))
                        {
                            result.ColumnData[c].Add(numVal);
                        }
                        else if (string.IsNullOrEmpty(cellText))
                        {
                            result.ColumnData[c].Add(null);
                        }
                        else
                        {
                            result.ColumnData[c].Add(cellText);
                            result.IsNumericColumn[c] = false;
                        }
                    }
                    if (allEmpty) break;
                }

                result.DataRowCount = result.ColumnData[0].Count;
            }
            catch (Exception ex)
            {
                result.Errors.Add($"读取 Excel 失败: {ex.Message}");
            }
            return result;
        }

        private class DirectImportRawData
        {
            public List<string> Headers { get; set; } = new();
            public List<List<object?>> ColumnData { get; set; } = new();
            public List<bool> IsNumericColumn { get; set; } = new();
            public int DataRowCount { get; set; }
            public List<string> Errors { get; set; } = new();
        }

        // ── 直接导入结果数据类 ──
        private class DirectImportResult
        {
            public List<string> FactorNames { get; set; } = new();
            public List<string> ResponseNames { get; set; } = new();
            public Dictionary<string, string> FactorTypes { get; set; } = new();
            public List<Dictionary<string, object>> FactorsData { get; set; } = new();
            public Dictionary<string, List<double>> ResponsesData { get; set; } = new();
            public int DataCount { get; set; }
            public List<string> Errors { get; set; } = new();
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
        private class ParetoResult
        {
            [JsonProperty("effects")] public List<ParetoEffectItem> Effects { get; set; } = new();
            [JsonProperty("log_worth_crit")] public double LogWorthCrit { get; set; } = 1.301;
            [JsonProperty("alpha")] public double Alpha { get; set; } = 0.05;
            [JsonProperty("measure")] public string Measure { get; set; } = "LogWorth";
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
            [JsonProperty("y_lower")] public List<double> YLower { get; set; } = new();
            [JsonProperty("y_upper")] public List<double> YUpper { get; set; } = new();
            /// <summary>连续因子的 X 值（double 列表）</summary>
            public List<double> X => IsCategorical ? new() : XRaw.Select(v => Convert.ToDouble(v)).ToList();
            /// <summary>类别因子的 X 标签</summary>
            public List<string> X_Labels => IsCategorical ? XRaw.Select(v => v?.ToString() ?? "").ToList() : new();
            /// <summary>连续因子的当前值</summary>
            public double CurrentNumericValue => IsCategorical ? 0 : Convert.ToDouble(CurrentValue ?? 0);
        }
        private class SurfaceData
        {
            [JsonProperty("x")] public List<double> X { get; set; } = new();
            [JsonProperty("y")] public List<double> Y { get; set; } = new();
            [JsonProperty("z")] public List<List<double>> Z { get; set; } = new();
            [JsonProperty("x_label")] public string XLabel { get; set; } = "";
            [JsonProperty("y_label")] public string YLabel { get; set; } = "";
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

        // ── 异常点 ──

        public class OutlierItem : BindableBase
        {
            public int Index { get; set; }
            public double Actual { get; set; }
            public double Predicted { get; set; }
            public double Residual { get; set; }
            public double StdResidual { get; set; }
            public double CooksD { get; set; }
            public double Leverage { get; set; }
            public string Reasons { get; set; } = "";

            private bool _isExcluded;
            public bool IsExcluded { get => _isExcluded; set => SetProperty(ref _isExcluded, value); }
        }

        // 反序列化用
        private class OutlierAnalysisResult
        {
            [JsonProperty("outliers")] public List<OutlierData> Outliers { get; set; } = new();
            [JsonProperty("thresholds")] public ThresholdData Thresholds { get; set; } = new();
            [JsonProperty("total_observations")] public int TotalObservations { get; set; }
            [JsonProperty("outlier_count")] public int OutlierCount { get; set; }
        }
        private class OutlierData
        {
            [JsonProperty("index")] public int Index { get; set; }
            [JsonProperty("actual")] public double Actual { get; set; }
            [JsonProperty("predicted")] public double Predicted { get; set; }
            [JsonProperty("residual")] public double Residual { get; set; }
            [JsonProperty("std_residual")] public double StdResidual { get; set; }
            [JsonProperty("cooks_d")] public double CooksD { get; set; }
            [JsonProperty("leverage")] public double Leverage { get; set; }
            [JsonProperty("reasons")] public List<string> Reasons { get; set; } = new();
        }
        private class ThresholdData
        {
            [JsonProperty("cooks_d")] public double CooksD { get; set; }
            [JsonProperty("leverage")] public double Leverage { get; set; }
            [JsonProperty("std_residual")] public double StdResidual { get; set; }
        }

        // ── 主效应 / 交互效应 ──

        private class MainEffectPoint
        {
            [JsonProperty("level")] public object Level { get; set; } = 0.0;
            [JsonProperty("mean")] public double Mean { get; set; }
        }
        private class InteractionEffectData
        {
            [JsonProperty("factor1")] public string Factor1 { get; set; } = "";
            [JsonProperty("factor2")] public string Factor2 { get; set; } = "";
            [JsonProperty("data")] public List<InteractionPoint> Data { get; set; } = new();
        }
        private class InteractionPoint
        {
            [JsonProperty("f1")] public object F1 { get; set; } = 0.0;
            [JsonProperty("f2")] public object F2 { get; set; } = 0.0;
            [JsonProperty("mean")] public double Mean { get; set; }
        }

        // ── Tukey HSD ──

        public class TukeyComparisonItem
        {
            public string Group1 { get; set; } = "";
            public string Group2 { get; set; } = "";
            public double MeanDiff { get; set; }
            public double PValue { get; set; }
            public double CILower { get; set; }
            public double CIUpper { get; set; }
            public bool Significant { get; set; }

            public string ComparisonText => $"{Group1} vs {Group2}";
            public string DiffText => $"{MeanDiff:+0.0000;-0.0000}";
            public string PValueText => PValue < 0.001 ? "p<0.001" : $"p={PValue:F4}";
            public string CIText => $"[{CILower:F2}, {CIUpper:F2}]";
            public string SignificanceMarker => Significant ? "★" : "";
        }

        // 反序列化用
        private class TukeyHSDResult
        {
            [JsonProperty("factor_name")] public string FactorName { get; set; } = "";
            [JsonProperty("comparisons")] public List<TukeyComparison> Comparisons { get; set; } = new();
            [JsonProperty("group_means")] public Dictionary<string, double> GroupMeans { get; set; } = new();
            [JsonProperty("mse")] public double MSE { get; set; }
            [JsonProperty("alpha")] public double Alpha { get; set; }
            [JsonProperty("error")] public string? Error { get; set; }
        }
        private class TukeyComparison
        {
            [JsonProperty("group1")] public string Group1 { get; set; } = "";
            [JsonProperty("group2")] public string Group2 { get; set; } = "";
            [JsonProperty("mean_diff")] public double MeanDiff { get; set; }
            [JsonProperty("p_value")] public double PValue { get; set; }
            [JsonProperty("ci_lower")] public double CILower { get; set; }
            [JsonProperty("ci_upper")] public double CIUpper { get; set; }
            [JsonProperty("significant")] public bool Significant { get; set; }
        }

        private class BoxCoxResult
        {
            [JsonProperty("optimal_lambda")] public double OptimalLambda { get; set; }
            [JsonProperty("lambda_ci")] public List<double>? LambdaCI { get; set; }
            [JsonProperty("rounded_lambda")] public double RoundedLambda { get; set; }
            [JsonProperty("transform_name")] public string TransformName { get; set; } = "";
            [JsonProperty("original_r_squared")] public double OriginalRSquared { get; set; }
            [JsonProperty("transformed_r_squared")] public double TransformedRSquared { get; set; }
            [JsonProperty("original_r_squared_adj")] public double OriginalRSquaredAdj { get; set; }
            [JsonProperty("transformed_r_squared_adj")] public double TransformedRSquaredAdj { get; set; }
            [JsonProperty("improvement")] public double Improvement { get; set; }
            [JsonProperty("recommend_transform")] public bool RecommendTransform { get; set; }
            [JsonProperty("recommendation")] public string Recommendation { get; set; } = "";
            [JsonProperty("lambda_profile")] public LambdaProfileData? LambdaProfile { get; set; }
            [JsonProperty("all_positive")] public bool AllPositive { get; set; }
            [JsonProperty("error")] public string? Error { get; set; }
        }

        private class LambdaProfileData
        {
            [JsonProperty("lambdas")] public List<double> Lambdas { get; set; } = new();
            [JsonProperty("log_likelihoods")] public List<double?> LogLikelihoods { get; set; } = new();
        }
    }


}