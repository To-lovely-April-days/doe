using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using MaxChemical.Infrastructure.DOE;
using MaxChemical.Infrastructure.Services;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Events;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using Newtonsoft.Json;
using Prism.Commands;
using Prism.Events;
using Prism.Mvvm;
using System.Data;

namespace MaxChemical.Modules.DOE.ViewModels
{
    /// <summary>
    /// DOE 设计向导 ViewModel — 多步骤表单：
    /// Step1: 因子定义  Step2: 设计方法  Step3: 参数矩阵预览
    /// Step4: 响应变量 + 停止条件  Step5: 历史数据导入（可选）  Step6: 确认保存
    /// </summary>
    public class DOEDesignWizardViewModel : BindableBase
    {
        private readonly IDOEDesignService _designService;
        private readonly IDOERepository _repository;
        private readonly IGPRModelService _gprService;           //  新增: GPR 模型服务
        private readonly IMultiResponseGPRService _multiGprService;
        private readonly IFlowParameterProvider _paramProvider;
        private readonly IDialogService _dialogService;
        private readonly IEventAggregator _eventAggregator;      // ★ 新增: 事件聚合器
        private readonly ILogService _logger;

        // ── 向导状态 ──
        private int _currentStep = 1;
        private const int TOTAL_STEPS = 6;
        private bool _isLoading;
        private string _statusMessage = "";
        private string _batchName = "";

        // ── 因子数据 (Step1) ──
        private ObservableCollection<FactorViewModel> _factors = new();
        private ObservableCollection<FactorCandidate> _availableCandidates = new();

        // ── 设计方法 (Step2) ──
        private DOEDesignMethod _selectedMethod = DOEDesignMethod.FullFactorial;
        private string _taguchiTableType = "auto";
        private int _fractionalResolution = 3;
        //  新增: CCD 配置
        private string _ccdAlphaType = "rotatable";
        private int _ccdCenterCount = -1;
        //  新增: BBD 配置
        private int _bbdCenterCount = -1;
        //  新增: D-Optimal 配置
        private int _dOptimalRunCount = -1;
        private string _dOptimalModelType = "quadratic";
        //  新增: 设计质量
        private DOEDesignQuality? _designQuality;
        // ── 参数矩阵 (Step3) ──
        private DOEDesignMatrix? _designMatrix;
        private DataTable? _matrixTable;
        private int _centerPointCount = 0;
        private bool _isRandomized;

        // ── 响应变量 + 停止条件 (Step4) ──
        private ObservableCollection<ResponseViewModel> _responses = new();
        private ObservableCollection<StopConditionViewModel> _stopConditions = new();

        // ── 历史数据导入 (Step5) ──
        private string? _importFilePath;
        private DataValidationResult? _validationResult;
        private int _importedDataCount;
        private string? _customImportPath;  // ★ v6: 自定义导入 Excel 路径

        // ★ 新增: 项目模式字段
        private string? _currentProjectId;
        private int? _currentRoundNumber;
        private DOEProjectPhase? _currentProjectPhase;
        private bool _isProjectMode;

        public DOEDesignWizardViewModel(
            IDOEDesignService designService,
            IDOERepository repository,
            IGPRModelService gprService,                         //  新增
            IMultiResponseGPRService multiGprService,
            IFlowParameterProvider paramProvider,
            IDialogService dialogService,
            IEventAggregator eventAggregator,                    // ★ 新增
            ILogService logger)
        {
            _designService = designService;
            _repository = repository;
            _gprService = gprService;                            //  新增
            _paramProvider = paramProvider;
            _dialogService = dialogService;
            _eventAggregator = eventAggregator;                  // ★ 新增
            _multiGprService = multiGprService ?? throw new ArgumentNullException(nameof(multiGprService));
            _logger = logger?.ForContext<DOEDesignWizardViewModel>()
                      ?? throw new ArgumentNullException(nameof(logger));

            InitializeCommands();
            LoadAvailableCandidates();

            // ★ 新增: 订阅下一轮请求事件
            _eventAggregator.GetEvent<RequestNextRoundEvent>().Subscribe(payload =>
            {
                InitializeFromProject(payload);
            }, ThreadOption.UIThread);
        }

        // ══════════════ Properties ══════════════

        public int CurrentStep { get => _currentStep; set { SetProperty(ref _currentStep, value); UpdateStepVisibility(); } }
        public bool IsLoading { get => _isLoading; set => SetProperty(ref _isLoading, value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public string BatchName { get => _batchName; set => SetProperty(ref _batchName, value); }
        //  新增: CCD 属性
        public string CcdAlphaType { get => _ccdAlphaType; set => SetProperty(ref _ccdAlphaType, value); }
        public int CcdCenterCount { get => _ccdCenterCount; set => SetProperty(ref _ccdCenterCount, value); }

        //  新增: BBD 属性
        public int BbdCenterCount { get => _bbdCenterCount; set => SetProperty(ref _bbdCenterCount, value); }
        //
        //  新增: D-Optimal 属性
        public int DOptimalRunCount { get => _dOptimalRunCount; set => SetProperty(ref _dOptimalRunCount, value); }
        public string DOptimalModelType { get => _dOptimalModelType; set => SetProperty(ref _dOptimalModelType, value); }
        //  新增: 设计质量评估结果
        public DOEDesignQuality? DesignQuality
        {
            get => _designQuality; set => SetProperty(ref _designQuality, value);
        }
        //  新增: 设计质量显示属性（供 XAML 绑定）
        public string QualityDEffText => DesignQuality != null ? $"{DesignQuality.DEfficiency:P1}" : "—";
        public string QualityAEffText => DesignQuality != null ? $"{DesignQuality.AEfficiency:P1}" : "—";
        public string QualityGEffText => DesignQuality != null ? $"{DesignQuality.GEfficiency:P1}" : "—";
        public string QualityDofText => DesignQuality != null ? $"{DesignQuality.DegreesOfFreedom}" : "—";
        public bool HasDesignQuality => DesignQuality != null;

        // ★ 新增: 项目模式属性
        /// <summary>是否在项目模式下创建批次</summary>
        public bool IsProjectMode
        {
            get => _isProjectMode;
            set => SetProperty(ref _isProjectMode, value);
        }

        private string _projectName = "";
        /// <summary>项目名称（UI 显示用）</summary>
        public string ProjectName
        {
            get => _projectName;
            set => SetProperty(ref _projectName, value);
        }

        // Step1
        public ObservableCollection<FactorViewModel> Factors { get => _factors; set => SetProperty(ref _factors, value); }
        public ObservableCollection<FactorCandidate> AvailableCandidates { get => _availableCandidates; set => SetProperty(ref _availableCandidates, value); }

        // Step2
        public DOEDesignMethod SelectedMethod { get => _selectedMethod; set => SetProperty(ref _selectedMethod, value); }
        public string TaguchiTableType { get => _taguchiTableType; set => SetProperty(ref _taguchiTableType, value); }
        public int FractionalResolution { get => _fractionalResolution; set => SetProperty(ref _fractionalResolution, value); }

        // Step3
        public DataTable? MatrixTable { get => _matrixTable; set => SetProperty(ref _matrixTable, value); }
        public int MatrixRunCount => _designMatrix?.RunCount ?? 0;
        public int CenterPointCount { get => _centerPointCount; set => SetProperty(ref _centerPointCount, value); }
        public bool IsRandomized { get => _isRandomized; set => SetProperty(ref _isRandomized, value); }

        // Step4
        public ObservableCollection<ResponseViewModel> Responses { get => _responses; set => SetProperty(ref _responses, value); }
        public ObservableCollection<StopConditionViewModel> StopConditions { get => _stopConditions; set => SetProperty(ref _stopConditions, value); }

        // Step5
        public string? ImportFilePath
        {
            get => _importFilePath;
            set
            {
                if (SetProperty(ref _importFilePath, value))
                {
                    ValidateImportCommand.RaiseCanExecuteChanged();
                }
            }
        }
        public DataValidationResult? ValidationResult { get => _validationResult; set => SetProperty(ref _validationResult, value); }
        public int ImportedDataCount { get => _importedDataCount; set => SetProperty(ref _importedDataCount, value); }

        // Step visibility
        public bool IsStep1Visible => CurrentStep == 1;
        public bool IsStep2Visible => CurrentStep == 2;
        public bool IsStep3Visible => CurrentStep == 3;
        public bool IsStep4Visible => CurrentStep == 4;
        public bool IsStep5Visible => CurrentStep == 5;
        public bool IsStep6Visible => CurrentStep == 6;

        // ══════════════ Commands ══════════════

        public DelegateCommand NextStepCommand { get; private set; } = null!;
        public DelegateCommand PreviousStepCommand { get; private set; } = null!;
        public DelegateCommand AddFactorCommand { get; private set; } = null!;
        public DelegateCommand<FactorViewModel> RemoveFactorCommand { get; private set; } = null!;
        public DelegateCommand AddResponseCommand { get; private set; } = null!;
        public DelegateCommand<ResponseViewModel> RemoveResponseCommand { get; private set; } = null!;
        public DelegateCommand AddStopConditionCommand { get; private set; } = null!;
        public DelegateCommand<StopConditionViewModel> RemoveStopConditionCommand { get; private set; } = null!;
        public DelegateCommand GenerateMatrixCommand { get; private set; } = null!;
        public DelegateCommand RandomizeMatrixCommand { get; private set; } = null!;
        public DelegateCommand AddCenterPointsCommand { get; private set; } = null!;
        public DelegateCommand BrowseImportFileCommand { get; private set; } = null!;
        public DelegateCommand ValidateImportCommand { get; private set; } = null!;
        public DelegateCommand SaveBatchCommand { get; private set; } = null!;

        /// <summary>
        /// 保存完成事件 — View 层订阅此事件关闭窗口
        /// </summary>
        public event EventHandler<string>? BatchSaved;

        private void InitializeCommands()
        {
            NextStepCommand = new DelegateCommand(async () => await GoNextStepAsync(), () => CurrentStep < TOTAL_STEPS);
            PreviousStepCommand = new DelegateCommand(() => { CurrentStep--; }, () => CurrentStep > 1);
            AddFactorCommand = new DelegateCommand(AddFactor);
            RemoveFactorCommand = new DelegateCommand<FactorViewModel>(f => { if (f != null) Factors.Remove(f); });
            AddResponseCommand = new DelegateCommand(AddResponse);
            RemoveResponseCommand = new DelegateCommand<ResponseViewModel>(r => { if (r != null) Responses.Remove(r); });
            AddStopConditionCommand = new DelegateCommand(AddStopCondition);
            RemoveStopConditionCommand = new DelegateCommand<StopConditionViewModel>(c => { if (c != null) StopConditions.Remove(c); });
            GenerateMatrixCommand = new DelegateCommand(async () => await GenerateMatrixAsync());
            RandomizeMatrixCommand = new DelegateCommand(RandomizeMatrix, () => _designMatrix != null);
            AddCenterPointsCommand = new DelegateCommand(async () => await AddCenterPointsAsync(), () => _designMatrix != null);
            BrowseImportFileCommand = new DelegateCommand(BrowseImportFile);
            ValidateImportCommand = new DelegateCommand(async () => await ValidateImportAsync(), () => !string.IsNullOrEmpty(ImportFilePath));
            SaveBatchCommand = new DelegateCommand(async () => await SaveBatchAsync());
        }

        // ══════════════ ★ 新增: 项目模式初始化 ══════════════

        /// <summary>
        /// ★ 新增: 从项目模式初始化设计向导
        ///
        /// 由 RequestNextRoundEvent 触发调用。
        /// 自动预填活跃因子、推荐设计方法、关联项目 ID。
        /// </summary>
        public void InitializeFromProject(NextRoundPayload payload)
        {
            _currentProjectId = payload.ProjectId;
            _currentRoundNumber = payload.RoundNumber;
            _currentProjectPhase = payload.RecommendedPhase;
            IsProjectMode = true;

            // 预填因子
            Factors.Clear();
            foreach (var f in payload.PrefilledFactors)
            {
                var fvm = new FactorViewModel
                {
                    FactorName = f.FactorName,
                    FactorType = f.FactorType,
                    LowerBound = f.LowerBound,
                    UpperBound = f.UpperBound,
                    LevelCount = f.LevelCount,
                    CategoryLevels = f.CategoryLevels ?? "",
                    AvailableCandidates = new ObservableCollection<FactorCandidate>(AvailableCandidates)
                };
                // 尝试绑定参数候选
                if (!string.IsNullOrEmpty(f.SourceNodeId) && !string.IsNullOrEmpty(f.SourceParamName))
                {
                    var candidate = fvm.AvailableCandidates
                        .FirstOrDefault(c => c.NodeId == f.SourceNodeId && c.ParameterName == f.SourceParamName);
                    if (candidate != null)
                        fvm.SelectedCandidate = candidate;
                }
                Factors.Add(fvm);
            }

            // 预选设计方法
            SelectedMethod = payload.RecommendedMethod;

            // ★ 优化2: 预填响应变量（从上一轮继承）
            if (payload.PrefilledResponses.Count > 0)
            {
                Responses.Clear();
                foreach (var r in payload.PrefilledResponses)
                {
                    Responses.Add(new ResponseViewModel
                    {
                        ResponseName = r.ResponseName,
                        Unit = r.Unit
                    });
                }
            }

            // 提示信息
            StatusMessage = $"项目模式: 第 {payload.RoundNumber} 轮 ({payload.RecommendedPhase})";
        }

        /// <summary>
        /// 将 DOEDesignMatrix 转为 DataTable（WPF DataGrid 原生支持）
        /// ★ 修复 (v3): 类别因子列用 string 类型显示标签
        /// </summary>
        private DataTable ConvertMatrixToDataTable(DOEDesignMatrix matrix)
        {
            var dt = new DataTable();

            // 添加序号列
            dt.Columns.Add("组号", typeof(int));

            // ★ 修复 (v3): 类别因子列用 string 类型
            var categoricalFactorNames = new HashSet<string>(
                Factors.Where(f => f.IsCategorical).Select(f => f.FactorName));

            foreach (var name in matrix.FactorNames)
            {
                if (categoricalFactorNames.Contains(name))
                    dt.Columns.Add(name, typeof(string));
                else
                    dt.Columns.Add(name, typeof(double));
            }

            // 填充数据
            for (int i = 0; i < matrix.Rows.Count; i++)
            {
                var dr = dt.NewRow();
                dr["组号"] = i + 1;
                foreach (var name in matrix.FactorNames)
                {
                    if (matrix.Rows[i].TryGetValue(name, out var val))
                    {
                        if (categoricalFactorNames.Contains(name))
                            dr[name] = val?.ToString() ?? "";
                        else
                            dr[name] = val is double d ? d : Convert.ToDouble(val);
                    }
                    else
                    {
                        dr[name] = categoricalFactorNames.Contains(name) ? (object)"" : (object)0.0;
                    }
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }
        // ══════════════ Step Navigation ══════════════

        private async Task GoNextStepAsync()
        {
            // 每步切换前做校验
            switch (CurrentStep)
            {
                case 1:
                    if (Factors.Count == 0)
                    {
                        _dialogService.ShowError("请至少添加一个因子", "验证");
                        return;
                    }
                    foreach (var f in Factors)
                    {
                        // ★ 修改: 类别因子不需要绑定流程参数（手动操作）
                        if (f.IsContinuous && f.SelectedCandidate == null)
                        {
                            _dialogService.ShowError($"连续因子 '{f.FactorName}' 未绑定流程参数，请在下拉框中选择", "验证");
                            return;
                        }
                        // ★ 修改: 仅连续因子检查上下界
                        if (f.IsContinuous && f.LowerBound >= f.UpperBound)
                        {
                            _dialogService.ShowError($"因子 '{f.FactorName}' 的下界必须小于上界", "验证");
                            return;
                        }
                        // ★ 新增: 类别因子必须输入至少 2 个水平
                        if (f.IsCategorical)
                        {
                            var levels = (f.CategoryLevels ?? "").Split(',')
                                .Select(s => s.Trim()).Where(s => !string.IsNullOrEmpty(s)).ToList();
                            if (levels.Count < 2)
                            {
                                _dialogService.ShowError($"类别因子 '{f.FactorName}' 至少需要 2 个水平标签（逗号分隔）", "验证");
                                return;
                            }
                        }
                    }
                    break;

                case 2: // 设计方法 → 生成矩阵
                    //  修改: 连续因子数校验（类别因子不计入 RSM 最低要求）
                    var continuousCount = Factors.Count(f => f.IsContinuous);
                    if (SelectedMethod == DOEDesignMethod.BoxBehnken && continuousCount < 3)
                    {
                        _dialogService.ShowError("Box-Behnken 设计至少需要 3 个连续因子", "验证");
                        return;
                    }
                    if (SelectedMethod == DOEDesignMethod.CCD && continuousCount < 2)
                    {
                        _dialogService.ShowError("CCD 设计至少需要 2 个连续因子", "验证");
                        return;
                    }
                    // ★ 修复 v6: 自定义导入先选文件再导入矩阵
                    if (SelectedMethod == DOEDesignMethod.Custom)
                    {
                        var path = _dialogService.ShowOpenFileDialog("Excel 文件|*.xlsx");
                        if (string.IsNullOrEmpty(path))
                            return;
                        _customImportPath = path;
                    }
                    await GenerateMatrixAsync();
                    if (_designMatrix == null || _designMatrix.IsEmpty)
                    {
                        _dialogService.ShowError("矩阵生成失败，请检查因子配置或导入文件", "验证");
                        return;
                    }
                    break;

                case 4: // 响应变量 → 历史导入
                    if (Responses.Count == 0)
                    {
                        _dialogService.ShowError("请至少添加一个响应变量", "验证");
                        return;
                    }
                    break;
            }

            CurrentStep++;
        }

        private void UpdateStepVisibility()
        {
            RaisePropertyChanged(nameof(IsStep1Visible));
            RaisePropertyChanged(nameof(IsStep2Visible));
            RaisePropertyChanged(nameof(IsStep3Visible));
            RaisePropertyChanged(nameof(IsStep4Visible));
            RaisePropertyChanged(nameof(IsStep5Visible));
            RaisePropertyChanged(nameof(IsStep6Visible));
            NextStepCommand.RaiseCanExecuteChanged();
            PreviousStepCommand.RaiseCanExecuteChanged();
        }

        // ══════════════ Step1: 因子管理 ══════════════

        private void LoadAvailableCandidates()
        {
            try
            {
                var candidates = _paramProvider.GetFactorCandidates();
                AvailableCandidates = new ObservableCollection<FactorCandidate>(candidates);
                _logger.LogInformation("加载到 {Count} 个因子候选项", candidates.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载因子候选列表失败");
            }
        }

        private void AddFactor()
        {
            var factor = new FactorViewModel
            {
                FactorName = $"因子{Factors.Count + 1}",
                LowerBound = 0,
                UpperBound = 100,
                LevelCount = 3,
                AvailableCandidates = new ObservableCollection<FactorCandidate>(AvailableCandidates)
            };
            Factors.Add(factor);
        }



        // ══════════════ Step2+3: 矩阵生成 ══════════════
        private async Task GenerateMatrixAsync()
        {
            try
            {
                IsLoading = true;
                StatusMessage = "正在生成参数矩阵...";

                var factors = Factors.Select(f => f.ToModel()).ToList();

                //  修改: switch 新增 CCD / BoxBehnken / DOptimal / Custom 分支
                _designMatrix = SelectedMethod switch
                {
                    DOEDesignMethod.FullFactorial => await _designService.GenerateFullFactorialAsync(factors),
                    DOEDesignMethod.FractionalFactorial => await _designService.GenerateFractionalFactorialAsync(factors, FractionalResolution),
                    DOEDesignMethod.Taguchi => await _designService.GenerateTaguchiAsync(factors, TaguchiTableType),
                    DOEDesignMethod.CCD => await _designService.GenerateCCDAsync(factors, CcdAlphaType, CcdCenterCount),
                    DOEDesignMethod.BoxBehnken => await _designService.GenerateBoxBehnkenAsync(factors, BbdCenterCount),
                    DOEDesignMethod.DOptimal => await _designService.GenerateDOptimalAsync(factors, DOptimalRunCount, DOptimalModelType),
                    DOEDesignMethod.Custom => !string.IsNullOrEmpty(_customImportPath)
                        ? await _designService.ImportCustomMatrixAsync(_customImportPath, factors)
                        : null,
                    _ => null
                };

                if (_designMatrix != null)
                {
                    MatrixTable = ConvertMatrixToDataTable(_designMatrix);
                    RaisePropertyChanged(nameof(MatrixRunCount));
                    StatusMessage = $"已生成 {_designMatrix.RunCount} 组实验";

                    //  新增: 自动计算设计质量
                    await EvaluateDesignQualityAsync(factors);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成参数矩阵失败");
                _dialogService.ShowError($"生成失败: {ex.Message}", "错误");
            }
            finally
            {
                IsLoading = false;
                RandomizeMatrixCommand.RaiseCanExecuteChanged();
                AddCenterPointsCommand.RaiseCanExecuteChanged();
            }
        }
        /// <summary>
        ///  新增: 评估设计质量
        /// 在生成矩阵后自动调用，结果展示在 Step3 的设计质量面板
        /// </summary>
        private async Task EvaluateDesignQualityAsync(List<DOEFactor> factors)
        {
            if (_designMatrix == null || _designMatrix.IsEmpty)
            {
                DesignQuality = null;
                return;
            }

            try
            {
                StatusMessage = "正在评估设计质量...";

                // 根据设计方法选择合适的模型类型
                string modelType = SelectedMethod switch
                {
                    DOEDesignMethod.CCD or DOEDesignMethod.BoxBehnken or DOEDesignMethod.DOptimal => "quadratic",
                    DOEDesignMethod.FractionalFactorial => "linear",
                    _ => factors.Count <= 4 ? "interaction" : "linear"
                };

                DesignQuality = await _designService.GetDesignQualityAsync(factors, _designMatrix, modelType);

                // 通知 UI 更新
                RaisePropertyChanged(nameof(QualityDEffText));
                RaisePropertyChanged(nameof(QualityAEffText));
                RaisePropertyChanged(nameof(QualityGEffText));
                RaisePropertyChanged(nameof(QualityDofText));
                RaisePropertyChanged(nameof(HasDesignQuality));

                _logger.LogInformation(" 设计质量评估完成: D-eff={DEff}, DF={DF}",
                    DesignQuality.DEfficiency, DesignQuality.DegreesOfFreedom);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "设计质量评估失败");
                DesignQuality = null;
            }
        }
        private void RandomizeMatrix()
        {
            if (_designMatrix == null) return;
            _designMatrix.Randomize();
            MatrixTable = ConvertMatrixToDataTable(_designMatrix);
            IsRandomized = true;
        }

        private async Task AddCenterPointsAsync()
        {
            if (_designMatrix == null || CenterPointCount <= 0) return;

            var factors = Factors.Select(f => f.ToModel()).ToList();
            // ★ 修复 (v3): 类别因子中心值用第一个水平标签，连续因子用 CenterPoint
            var centerValues = new Dictionary<string, object>();
            foreach (var f in factors)
            {
                if (f.IsCategorical)
                {
                    var levels = f.GetCategoryLevelList();
                    centerValues[f.FactorName] = levels.Count > 0 ? levels[0] : "";
                }
                else
                {
                    centerValues[f.FactorName] = f.CenterPoint;
                }
            }

            for (int i = 0; i < CenterPointCount; i++)
                _designMatrix.AddCenterPoint(centerValues);

            MatrixTable = ConvertMatrixToDataTable(_designMatrix);
            RaisePropertyChanged(nameof(MatrixRunCount));
        }

        // ══════════════ Step4: 响应 + 停止条件 ══════════════

        private void AddResponse()
        {
            Responses.Add(new ResponseViewModel { ResponseName = $"响应{Responses.Count + 1}", Unit = "%" });
        }

        private void AddStopCondition()
        {
            StopConditions.Add(new StopConditionViewModel
            {
                ConditionType = DOEStopConditionType.Threshold,
                ResponseName = Responses.FirstOrDefault()?.ResponseName ?? "",
                Operator = "GreaterThanOrEqual",
                TargetValue = 95
            });
        }

        // ══════════════ Step5: 历史数据导入 ══════════════

        private void BrowseImportFile()
        {
            var path = _dialogService.ShowOpenFileDialog("Excel 文件|*.xlsx");
            if (!string.IsNullOrEmpty(path))
            {
                ImportFilePath = path;
                ValidateImportCommand.RaiseCanExecuteChanged();
            }
        }

        private async Task ValidateImportAsync()
        {
            if (string.IsNullOrEmpty(ImportFilePath)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "正在校验导入文件...";

                var factors = Factors.Select(f => f.ToModel()).ToList();
                var responses = Responses.Select(r => r.ToModel()).ToList();

                ValidationResult = await _designService.ValidateImportedDataAsync(ImportFilePath, factors, responses);

                if (ValidationResult.IsValid)
                    StatusMessage = $"校验通过，{ValidationResult.ValidRowCount} 行有效数据";
                else
                    StatusMessage = $"校验失败: {string.Join("; ", ValidationResult.Errors)}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "校验导入文件失败");
                _dialogService.ShowError($"校验失败: {ex.Message}", "错误");
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ══════════════ Step6: 保存批次 ══════════════

        private async Task SaveBatchAsync()
        {
            try
            {
                IsLoading = true;
                StatusMessage = "正在保存 DOE 方案...";

                if (string.IsNullOrWhiteSpace(BatchName))
                    BatchName = $"DOE_{DateTime.Now:yyyyMMdd_HHmmss}";

                // 创建批次
                var batch = new DOEBatch
                {
                    FlowId = _paramProvider.GetCurrentFlowId() ?? "unknown",
                    FlowName = _paramProvider.GetCurrentFlowName() ?? "未命名流程",
                    BatchName = BatchName,
                    DesignMethod = SelectedMethod,
                    Status = DOEBatchStatus.Ready,
                    DesignConfigJson = JsonConvert.SerializeObject(new
                    {
                        method = SelectedMethod.ToString(),
                        taguchiTable = TaguchiTableType,
                        fractionalResolution = FractionalResolution,
                        isRandomized = IsRandomized,
                        centerPointCount = CenterPointCount,
                        //  新增: RSM 参数
                        ccdAlphaType = CcdAlphaType,
                        ccdCenterCount = CcdCenterCount,
                        bbdCenterCount = BbdCenterCount,
                        dOptimalRunCount = DOptimalRunCount,
                        dOptimalModelType = DOptimalModelType,
                        //  新增: 设计质量快照
                        designQuality = DesignQuality != null ? new
                        {
                            dEfficiency = DesignQuality.DEfficiency,
                            aEfficiency = DesignQuality.AEfficiency,
                            gEfficiency = DesignQuality.GEfficiency,
                            degreesOfFreedom = DesignQuality.DegreesOfFreedom
                        } : null
                    }),
                    // ★ 新增: 项目关联（_currentProjectId 为 null 时行为与旧版一致）
                    ProjectId = _currentProjectId,
                    RoundNumber = _currentRoundNumber,
                    ProjectPhase = _currentProjectPhase
                };

                var batchId = await _repository.CreateBatchAsync(batch);

                // 保存因子
                var factors = Factors.Select(f => f.ToModel()).ToList();
                await _repository.SaveFactorsAsync(batchId, factors);

                // 保存响应变量
                var responses = Responses.Select(r => r.ToModel()).ToList();
                await _repository.SaveResponsesAsync(batchId, responses);

                // 保存停止条件
                var stopConditions = StopConditions.Select(c => c.ToModel()).ToList();
                await _repository.SaveStopConditionsAsync(batchId, stopConditions);
                //  新增: 保存 Desirability 配置
                var desirabilityConfigs = Responses.Select(r => r.ToDesirabilityConfig()).ToList();
                var configJson = JsonConvert.SerializeObject(desirabilityConfigs);
                await _repository.SaveDesirabilityConfigAsync(batchId, configJson);
                _logger.LogInformation("Desirability 配置已保存: {Count} 个响应", desirabilityConfigs.Count);
                // 保存参数矩阵为 DOE runs (Pending 状态)
                if (_designMatrix != null)
                {
                    var runs = new List<DOERunRecord>();
                    for (int i = 0; i < _designMatrix.RunCount; i++)
                    {
                        runs.Add(new DOERunRecord
                        {
                            BatchId = batchId,
                            RunIndex = i,
                            FactorValuesJson = JsonConvert.SerializeObject(_designMatrix.Rows[i]),
                            DataSource = DOEDataSource.Measured,
                            Status = DOERunStatus.Pending
                        });
                    }
                    await _repository.SaveRunsAsync(runs);
                }

                // 导入历史数据（如果有）
                if (ValidationResult?.IsValid == true && !string.IsNullOrEmpty(ImportFilePath))
                {
                    var importedRuns = await _designService.ImportHistoricalDataAsync(
                        ImportFilePath, batchId, factors, responses);
                    await _repository.SaveRunsAsync(importedRuns);
                    ImportedDataCount = importedRuns.Count;
                }

                // ════════════════════════════════════════════════════════
                //  新增: 创建并持久化 GPR 模型（错误不阻断方案保存）
                // ════════════════════════════════════════════════════════
                await InitializeGPRModelAsync(batch.FlowId, batchId, factors, responses);

                // ★ 新增: 项目模式下更新项目信息 + 同步因子池
                if (!string.IsNullOrEmpty(_currentProjectId))
                {
                    try
                    {
                        var project = await _repository.GetProjectAsync(_currentProjectId);
                        if (project != null)
                        {
                            if (_currentProjectPhase.HasValue)
                                project.CurrentPhase = _currentProjectPhase.Value;

                            // ★ 首轮时关联 FlowId
                            if (string.IsNullOrEmpty(project.FlowId))
                            {
                                project.FlowId = batch.FlowId;
                                project.FlowName = batch.FlowName;
                            }

                            await _repository.UpdateProjectAsync(project);

                            // ★ 首轮时同步因子到项目因子池
                            var existingFactors = await _repository.GetProjectFactorsAsync(_currentProjectId);
                            if (existingFactors.Count == 0)
                            {
                                var projectFactors = factors.Select((f, idx) => new DOEProjectFactor
                                {
                                    ProjectId = _currentProjectId,
                                    FactorName = f.FactorName,
                                    FactorType = f.FactorType,
                                    FactorStatus = ProjectFactorStatus.Active,
                                    CurrentLowerBound = f.LowerBound,
                                    CurrentUpperBound = f.UpperBound,
                                    CategoryLevels = f.CategoryLevels,
                                    SourceNodeId = f.SourceNodeId,
                                    SourceParamName = f.SourceParamName,
                                    BoundsHistoryJson = JsonConvert.SerializeObject(new[]
                                    {
                                        new { batch_id = batchId, lower = f.LowerBound, upper = f.UpperBound,
                                              reason = "初始范围", timestamp = DateTime.Now }
                                    }),
                                    SortOrder = idx
                                }).ToList();

                                await _repository.SaveProjectFactorsAsync(_currentProjectId, projectFactors);
                                _logger.LogInformation("项目因子池已初始化: {Count} 个因子", projectFactors.Count);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "更新项目信息失败（不影响方案保存）");
                    }
                }

                StatusMessage = $"DOE 方案已保存: {batchId}";
                _logger.LogInformation("DOE 方案保存成功: {BatchId} - {BatchName}", batchId, BatchName);

                BatchSaved?.Invoke(this, batchId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存 DOE 方案失败");
                _dialogService.ShowError($"保存失败: {ex.Message}", "错误");
            }
            finally
            {
                IsLoading = false;
            }
        }

        /// <summary>
        ///  新增: 在创建 DOE 方案时初始化 GPR 模型
        /// 
        /// 执行流程：
        /// 1. InitializeModelAsync — 按 FlowId + FactorSignature 查已有模型或新建
        /// 2. 如果有导入的历史数据 → AppendData 逐条喂给模型
        /// 3. SaveInitialStateAsync — 持久化（有数据时自动训练）
        /// 
        /// 错误隔离: GPR 任何步骤失败都不影响方案保存
        /// </summary>
        private async Task InitializeGPRModelAsync(
              string flowId,
              string batchId,
              List<DOEFactor> factors,
              List<DOEResponse> responses)
        {
            try
            {
                StatusMessage = "正在初始化 GPR 预测模型...";
                // ★ 新增：传入项目 ID
                _gprService.SetProjectId(_currentProjectId);
                // Step 1: 初始化主响应模型（新建或恢复已有同签名模型）
                await _gprService.InitializeModelAsync(flowId, factors);

                //  新增 Step 1.5: 初始化多响应模型（为每个响应创建独立的 GPR）
                try
                {
                    await _multiGprService.InitializeAsync(flowId, factors, responses, _currentProjectId);
                    _logger.LogInformation("多响应 GPR 初始化成功: {Count} 个响应", responses.Count);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "多响应 GPR 初始化失败（不影响主模型）");
                }

                // Step 2: 如果有已导入的历史数据，喂给所有模型
                if (ImportedDataCount > 0)
                {
                    var primaryResponseName = responses.FirstOrDefault()?.ResponseName;
                    if (!string.IsNullOrEmpty(primaryResponseName))
                    {
                        var completedRuns = await _repository.GetCompletedRunsAsync(batchId);
                        int fedCount = 0;

                        foreach (var run in completedRuns)
                        {
                            try
                            {
                                // ★ 修复 (v3): 用 Dict<string, object> 反序列化，保留类别因子标签
                                var factorValues = JsonConvert.DeserializeObject<Dictionary<string, object>>(
                                    run.FactorValuesJson);
                                if (factorValues == null) continue;

                                var responseValues = !string.IsNullOrEmpty(run.ResponseValuesJson)
                                    ? JsonConvert.DeserializeObject<Dictionary<string, double>>(run.ResponseValuesJson)
                                    : null;
                                if (responseValues == null) continue;

                                //  修改: 喂给所有响应模型（原来只喂主响应）
                                _multiGprService.AppendAllResponses(
                                    factorValues, responseValues,
                                    "imported", "", "");

                                fedCount++;
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError(ex, "GPR 导入数据第 {RunIndex} 组失败，跳过", run.RunIndex);
                            }
                        }

                        _logger.LogInformation("GPR 历史数据导入: {FedCount}/{ImportedCount} 组",
                            fedCount, ImportedDataCount);
                    }
                }

                // Step 3: 持久化主模型
                await _gprService.SaveInitialStateAsync(flowId);

                //  新增 Step 3.5: 持久化次要模型
                try
                {
                    await _multiGprService.SaveAllAsync(flowId);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "次要 GPR 模型保存失败（不影响方案保存）");
                }

                _logger.LogInformation("GPR 模型已随方案创建: FlowId={FlowId}, DataCount={Count}",
                    flowId, _gprService.DataCount);
            }
            catch (Exception ex)
            {
                //  关键: GPR 失败不阻断方案保存
                _logger.LogError(ex, "GPR 模型初始化失败（不影响 DOE 方案保存）");
            }
        }
    }

    // ══════════════ 子 ViewModel ══════════════

    public class FactorViewModel : BindableBase
    {
        private string _factorName = "";
        private FactorCandidate? _selectedCandidate;
        private double _lowerBound;
        private double _upperBound;
        private int _levelCount = 3;
        private DOEFactorType _factorType = DOEFactorType.Continuous;
        private string _categoryLevels = "";

        /// <summary>
        /// 用户自定义的因子名称（如 "反应温度"）
        /// </summary>
        public string FactorName { get => _factorName; set => SetProperty(ref _factorName, value); }

        /// <summary>
        /// ★ 新增: 因子类型（连续/类别）
        /// </summary>
        public DOEFactorType FactorType
        {
            get => _factorType;
            set
            {
                if (SetProperty(ref _factorType, value))
                {
                    RaisePropertyChanged(nameof(IsContinuous));
                    RaisePropertyChanged(nameof(IsCategorical));
                    RaisePropertyChanged(nameof(StepSize));
                }
            }
        }

        /// <summary>★ 新增: 是否连续因子（UI 绑定用）</summary>
        public bool IsContinuous => FactorType == DOEFactorType.Continuous;

        /// <summary>★ 新增: 是否类别因子（UI 绑定用）</summary>
        public bool IsCategorical => FactorType == DOEFactorType.Categorical;

        /// <summary>
        /// ★ 新增: 类别因子的水平标签（逗号分隔，如 "催化剂A,催化剂B,催化剂C"）
        /// </summary>
        public string CategoryLevels
        {
            get => _categoryLevels;
            set
            {
                if (SetProperty(ref _categoryLevels, value))
                {
                    // 自动更新 LevelCount
                    var levels = (value ?? "").Split(',')
                        .Select(s => s.Trim())
                        .Where(s => !string.IsNullOrEmpty(s))
                        .ToList();
                    LevelCount = Math.Max(levels.Count, 1);
                }
            }
        }

        /// <summary>★ 新增: 因子类型选项（供 ComboBox 绑定）</summary>
        public static List<DOEFactorType> FactorTypeOptions => new()
        {
            DOEFactorType.Continuous,
            DOEFactorType.Categorical
        };

        /// <summary>
        /// 绑定的流程参数候选项（下拉选择）
        /// </summary>
        public FactorCandidate? SelectedCandidate
        {
            get => _selectedCandidate;
            set
            {
                if (SetProperty(ref _selectedCandidate, value) && value != null)
                {
                    // 自动填充默认值
                    if (string.IsNullOrEmpty(FactorName) || FactorName.StartsWith("因子"))
                    {
                        // 取参数名作为因子名
                        FactorName = value.ParameterName ?? value.DisplayName;
                    }

                    // 根据当前值自动设置上下界（仅连续因子）
                    if (IsContinuous && value.CurrentValue != null && double.TryParse(value.CurrentValue.ToString(), out var v) && v != 0)
                    {
                        if (LowerBound == 0 && UpperBound == 100) // 只在默认值时自动填充
                        {
                            LowerBound = Math.Round(v * 0.5, 2);
                            UpperBound = Math.Round(v * 1.5, 2);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 绑定参数的显示文本
        /// </summary>
        public string BindingDisplayText => SelectedCandidate != null
            ? SelectedCandidate.DisplayName
            : "请选择绑定参数...";

        /// <summary>
        /// 是否已绑定
        /// </summary>
        public bool IsBound => SelectedCandidate != null;

        public double LowerBound { get => _lowerBound; set { SetProperty(ref _lowerBound, value); RaisePropertyChanged(nameof(StepSize)); } }
        public double UpperBound { get => _upperBound; set { SetProperty(ref _upperBound, value); RaisePropertyChanged(nameof(StepSize)); } }
        public int LevelCount { get => _levelCount; set { SetProperty(ref _levelCount, value); RaisePropertyChanged(nameof(StepSize)); } }
        public double StepSize => IsContinuous && LevelCount > 1 ? Math.Round((UpperBound - LowerBound) / (LevelCount - 1), 4) : 0;

        /// <summary>
        /// 可选的候选参数列表（绑定到 ComboBox 的 ItemsSource）
        /// </summary>
        public ObservableCollection<FactorCandidate> AvailableCandidates { get; set; } = new();

        public DOEFactor ToModel() => new()
        {
            FactorName = FactorName,
            FactorSource = SelectedCandidate?.SourceType ?? FactorSourceType.ParameterOverride,
            SourceNodeId = SelectedCandidate?.NodeId,
            SourceParamName = SelectedCandidate?.ParameterName,
            LowerBound = LowerBound,
            UpperBound = UpperBound,
            LevelCount = LevelCount,
            FactorType = FactorType,
            CategoryLevels = IsCategorical ? CategoryLevels : null
        };
    }

    public class ResponseViewModel : BindableBase
    {
        private string _responseName = "";
        private string _unit = "";
        //  新增: Desirability 配置（Step4 中设置）
        private DesirabilityGoal _goal = DesirabilityGoal.Maximize;
        private double _desirabilityLower = 0;
        private double _desirabilityUpper = 100;
        private double _desirabilityTarget = 100;
        private double _desirabilityWeight = 1.0;
        private int _importance = 3;
        public string ResponseName { get => _responseName; set => SetProperty(ref _responseName, value); }
        public string Unit { get => _unit; set => SetProperty(ref _unit, value); }
        //  新增: Desirability 属性
        public DesirabilityGoal Goal { get => _goal; set => SetProperty(ref _goal, value); }
        public double DesirabilityLower { get => _desirabilityLower; set => SetProperty(ref _desirabilityLower, value); }
        public double DesirabilityUpper { get => _desirabilityUpper; set => SetProperty(ref _desirabilityUpper, value); }
        public double DesirabilityTarget { get => _desirabilityTarget; set => SetProperty(ref _desirabilityTarget, value); }
        public double DesirabilityWeight { get => _desirabilityWeight; set => SetProperty(ref _desirabilityWeight, value); }
        public int Importance { get => _importance; set => SetProperty(ref _importance, value); }
        /// <summary>
        /// 目标类型选项（供 ComboBox 绑定）
        /// </summary>
        public static List<DesirabilityGoal> GoalOptions => new()
        {
            DesirabilityGoal.Maximize,
            DesirabilityGoal.Minimize,
            DesirabilityGoal.Target
        };
        public DOEResponse ToModel() => new()
        {
            ResponseName = ResponseName,
            Unit = Unit,
            CollectionMethod = DOECollectionMethod.Manual
        };
        /// <summary>
        ///  新增: 转为 Desirability 配置
        /// </summary>
        public DesirabilityResponseConfig ToDesirabilityConfig() => new()
        {
            ResponseName = ResponseName,
            Goal = Goal,
            Lower = DesirabilityLower,
            Upper = DesirabilityUpper,
            Target = DesirabilityTarget,
            Weight = DesirabilityWeight,
            Importance = Importance
        };
    }

    public class StopConditionViewModel : BindableBase
    {
        private DOEStopConditionType _conditionType = DOEStopConditionType.Threshold;
        private string _responseName = "";
        private string _operator = "GreaterThanOrEqual";
        private double _targetValue = 95;

        public DOEStopConditionType ConditionType { get => _conditionType; set => SetProperty(ref _conditionType, value); }
        public string ResponseName { get => _responseName; set => SetProperty(ref _responseName, value); }
        public string Operator { get => _operator; set => SetProperty(ref _operator, value); }
        public double TargetValue { get => _targetValue; set => SetProperty(ref _targetValue, value); }

        public static List<string> OperatorOptions => new()
        {
            "Equals", "NotEquals", "GreaterThan", "GreaterThanOrEqual", "LessThan", "LessThanOrEqual"
        };

        public DOEStopCondition ToModel() => new()
        {
            ConditionType = ConditionType,
            ResponseName = ResponseName,
            Operator = Operator,
            TargetValue = TargetValue
        };
    }
}