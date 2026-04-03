using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using MaxChemical.Infrastructure.Services;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using MaxChemical.Modules.DOE.Services;
using Newtonsoft.Json;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using Prism.Commands;
using Prism.Mvvm;

namespace MaxChemical.Modules.DOE.ViewModels
{
    /// <summary>
    /// DOE 分析视图 ViewModel — 管理所有图表的数据加载和 OxyPlot PlotModel 渲染
    ///  修改: 集成 ModelRouter + OLS 分析 + Desirability 多响应优化
    /// </summary>
    public class DOEAnalysisViewModel : BindableBase
    {
        private readonly IDOEAnalysisService _analysisService;
        private readonly IDOERepository _repository;
        private readonly ILogService _logger;
        private readonly IModelRouter _modelRouter;
        private readonly IDesirabilityService _desirabilityService;
        private readonly IDialogService _dialogService;          //  修复: 补全字段声明

        // ── 批次关联 ──
        private string _batchId = "";
        private string _currentBatchId = "";                     //  修复: 补全字段声明
        private string _currentFlowId = "";                      //  修复: 补全字段声明
        private bool _isLoading;
        private string _statusMessage = "";

        // ── 图表 PlotModels ──
        private PlotModel? _mainEffectsPlot;
        private PlotModel? _paretoPlot;
        private PlotModel? _actualVsPredictedPlot;
        private PlotModel? _residualPlot;
        private BitmapImage? _surfaceImage;

        // ── 文本输出 ──
        private string _anovaTableText = "";
        private string _regressionEquation = "";
        private string _rSquaredText = "";
        private string _modelEvolutionText = "";

        // ── 曲面图因子选择 ──
        private List<string> _factorNames = new();
        private List<string> _responseNames = new();             //  修复: 补全字段声明
        private string? _surfaceFactor1;
        private string? _surfaceFactor2;

        // ──  新增: OLS + ModelRouter + Desirability ──
        private AnalysisStrategy _currentStrategy = AnalysisStrategy.GPR;
        private string _strategyReason = "";
        private OLSAnalysisResult? _olsResult;
        private bool _isOlsMode;
        private string _selectedResponseName = "";
        private DesirabilityResult? _desirabilityResult;

        // ──  新增: 最优参数 ──
        private List<OptimalFactorItem> _optimalFactors = new();
        private bool _hasOptimalResult;

        public DOEAnalysisViewModel(
            IDOEAnalysisService analysisService,
            IDOERepository repository,
            IModelRouter modelRouter,
            IDesirabilityService desirabilityService,
            IDialogService dialogService,                        //  修复: 补全构造参数
            ILogService logger)
        {
            _analysisService = analysisService;
            _repository = repository;
            _logger = logger?.ForContext<DOEAnalysisViewModel>() ?? throw new ArgumentNullException(nameof(logger));
            _modelRouter = modelRouter;
            _desirabilityService = desirabilityService;
            _dialogService = dialogService;                      //  修复: 赋值

            RefreshAllCommand = new DelegateCommand(async () => await LoadAllChartsAsync(), () => !string.IsNullOrEmpty(BatchId));
            GenerateSurfaceCommand = new DelegateCommand(async () => await LoadSurfaceAsync(),
                () => !string.IsNullOrEmpty(SurfaceFactor1) && !string.IsNullOrEmpty(SurfaceFactor2) && SurfaceFactor1 != SurfaceFactor2);
            RunDesirabilityCommand = new DelegateCommand(async () => await RunDesirabilityOptimizationAsync());
        }

        // ══════════════ Properties ══════════════

        public string BatchId { get => _batchId; set { SetProperty(ref _batchId, value); RefreshAllCommand.RaiseCanExecuteChanged(); } }
        public bool IsLoading { get => _isLoading; set => SetProperty(ref _isLoading, value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        public PlotModel? MainEffectsPlot { get => _mainEffectsPlot; set => SetProperty(ref _mainEffectsPlot, value); }
        public PlotModel? ParetoPlot { get => _paretoPlot; set => SetProperty(ref _paretoPlot, value); }
        public PlotModel? ActualVsPredictedPlot { get => _actualVsPredictedPlot; set => SetProperty(ref _actualVsPredictedPlot, value); }
        public PlotModel? ResidualPlot { get => _residualPlot; set => SetProperty(ref _residualPlot, value); }
        public BitmapImage? SurfaceImage { get => _surfaceImage; set => SetProperty(ref _surfaceImage, value); }

        public string AnovaTableText { get => _anovaTableText; set => SetProperty(ref _anovaTableText, value); }
        public string RegressionEquation { get => _regressionEquation; set => SetProperty(ref _regressionEquation, value); }
        public string RSquaredText { get => _rSquaredText; set => SetProperty(ref _rSquaredText, value); }
        public string ModelEvolutionText { get => _modelEvolutionText; set => SetProperty(ref _modelEvolutionText, value); }

        public List<string> FactorNames { get => _factorNames; set => SetProperty(ref _factorNames, value); }
        public string? SurfaceFactor1 { get => _surfaceFactor1; set { SetProperty(ref _surfaceFactor1, value); GenerateSurfaceCommand.RaiseCanExecuteChanged(); } }
        public string? SurfaceFactor2 { get => _surfaceFactor2; set { SetProperty(ref _surfaceFactor2, value); GenerateSurfaceCommand.RaiseCanExecuteChanged(); } }

        // ──  新增属性 ──
        public AnalysisStrategy CurrentStrategy { get => _currentStrategy; set => SetProperty(ref _currentStrategy, value); }
        public string StrategyReason { get => _strategyReason; set => SetProperty(ref _strategyReason, value); }
        public bool IsOlsMode { get => _isOlsMode; set => SetProperty(ref _isOlsMode, value); }
        public bool IsGprMode => !IsOlsMode;
        public OLSAnalysisResult? OlsResult { get => _olsResult; set => SetProperty(ref _olsResult, value); }
        public string SelectedResponseName
        {
            get => _selectedResponseName;
            set { if (SetProperty(ref _selectedResponseName, value)) _ = RefreshOlsAnalysisAsync(); }
        }
        public List<string> ResponseNames { get => _responseNames; set => SetProperty(ref _responseNames, value); }
        public DesirabilityResult? DesirabilityResult { get => _desirabilityResult; set => SetProperty(ref _desirabilityResult, value); }
        public List<OptimalFactorItem> OptimalFactors { get => _optimalFactors; set => SetProperty(ref _optimalFactors, value); }
        public bool HasOptimalResult { get => _hasOptimalResult; set => SetProperty(ref _hasOptimalResult, value); }

        // ── Commands ──
        public DelegateCommand RefreshAllCommand { get; }
        public DelegateCommand GenerateSurfaceCommand { get; }
        public DelegateCommand RunDesirabilityCommand { get; }

        // ══════════════ 加载批次 ══════════════

        public async Task LoadBatchAsync(string batchId)
        {
            BatchId = batchId;
            _currentBatchId = batchId;
            var batch = await _repository.GetBatchWithDetailsAsync(batchId);
            if (batch == null) return;

            _currentFlowId = batch.FlowId;
            FactorNames = batch.Factors.Select(f => f.FactorName).ToList();
            ResponseNames = batch.Responses.Select(r => r.ResponseName).ToList();

            //  新增: 选择默认响应变量
            SelectedResponseName = ResponseNames.FirstOrDefault() ?? "";

            if (FactorNames.Count >= 2)
            {
                SurfaceFactor1 = FactorNames[0];
                SurfaceFactor2 = FactorNames[1];
            }

            //  新增: ModelRouter 自动选择分析策略
            var completedCount = batch.Runs.Count(r => r.Status == DOERunStatus.Completed);
            CurrentStrategy = _modelRouter.SelectStrategy(batch.DesignMethod, batch.Factors.Count, completedCount);
            StrategyReason = _modelRouter.GetStrategyReason(CurrentStrategy, batch.DesignMethod, batch.Factors.Count, completedCount);
            IsOlsMode = (CurrentStrategy == AnalysisStrategy.OLS);
            RaisePropertyChanged(nameof(IsGprMode));

            _logger.LogInformation(" ModelRouter 决策: Strategy={Strategy}, Method={Method}, Factors={K}, Data={N}",
                CurrentStrategy, batch.DesignMethod, batch.Factors.Count, completedCount);

            //  新增: 如果是 OLS 模式，加载 OLS 分析
            if (IsOlsMode && !string.IsNullOrEmpty(SelectedResponseName))
            {
                await RefreshOlsAnalysisAsync();
            }

            await LoadAllChartsAsync();
        }

        // ══════════════  新增: OLS 分析 ══════════════

        private async Task RefreshOlsAnalysisAsync()
        {
            if (!IsOlsMode || string.IsNullOrEmpty(_currentBatchId) || string.IsNullOrEmpty(SelectedResponseName))
                return;

            try
            {
                IsLoading = true;
                StatusMessage = "正在进行 OLS 回归分析...";

                OlsResult = await _analysisService.FitOlsAsync(_currentBatchId, SelectedResponseName, "quadratic");

                if (OlsResult?.ModelSummary != null)
                {
                    StatusMessage = $"OLS 分析完成: R²={OlsResult.ModelSummary.RSquared:F4}, " +
                                   $"R²adj={OlsResult.ModelSummary.RSquaredAdj:F4}, " +
                                   $"R²pred={OlsResult.ModelSummary.RSquaredPred:F4}";

                    _logger.LogInformation(" OLS 分析: R²={R2}, RMSE={RMSE}, LOF_P={LOFP}",
                        OlsResult.ModelSummary.RSquared,
                        OlsResult.ModelSummary.RMSE,
                        OlsResult.ModelSummary.LackOfFitP);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OLS 分析失败");
                StatusMessage = $"OLS 分析失败: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ══════════════  新增: Desirability 优化 ══════════════

        private async Task RunDesirabilityOptimizationAsync()
        {
            if (ResponseNames.Count < 2)
            {
                _dialogService.ShowError("Desirability 优化需要至少 2 个响应变量", "提示");
                return;
            }

            try
            {
                IsLoading = true;
                StatusMessage = "正在搜索多响应最优因子组合...";

                var configs = await _desirabilityService.LoadConfigAsync(_currentBatchId);
                if (configs.Count == 0)
                {
                    _dialogService.ShowError("请先在设计向导 Step4 中配置各响应变量的优化目标", "缺少配置");
                    return;
                }

                var batch = await _repository.GetBatchWithDetailsAsync(_currentBatchId);
                if (batch == null) return;

                await _desirabilityService.ConfigureAsync(_currentBatchId, configs, batch.Factors);
                DesirabilityResult = await _desirabilityService.OptimizeAsync();

                if (DesirabilityResult?.Success == true)
                {
                    StatusMessage = $"多响应优化完成: 综合 D = {DesirabilityResult.CompositeD:F4}";

                    if (DesirabilityResult.OptimalFactors.Count > 0)
                    {
                        // ★ 修复: OptimalFactors 现在是 Dict<string, object>
                        // 连续因子是 double，类别因子是 string
                        // BarWidth 只对连续因子有意义，类别因子 BarWidth=0
                        var items = DesirabilityResult.OptimalFactors.Select(kv => new OptimalFactorItem
                        {
                            Key = kv.Key,
                            Value = kv.Value
                        }).ToList();

                        // 计算连续因子的最大值用于 BarWidth 归一化
                        var numericValues = items.Where(x => x.IsNumeric).Select(x => Math.Abs(x.NumericValue)).ToList();
                        double maxVal = numericValues.Count > 0 ? numericValues.Max() : 1.0;

                        foreach (var item in items)
                        {
                            item.BarWidth = item.IsNumeric && maxVal > 0
                                ? Math.Abs(item.NumericValue) / maxVal * 200
                                : 0;
                        }

                        OptimalFactors = items;
                        HasOptimalResult = true;
                    }
                }
                else
                {
                    StatusMessage = $"优化失败: {DesirabilityResult?.ErrorMessage ?? "未知错误"}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Desirability 优化失败");
                _dialogService.ShowError($"优化失败: {ex.Message}", "错误");
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ══════════════ 加载所有图表 ══════════════

        private async Task LoadAllChartsAsync()
        {
            if (string.IsNullOrEmpty(BatchId)) return;

            try
            {
                IsLoading = true;
                StatusMessage = "正在加载分析图表...";

                await Task.WhenAll(
                    LoadMainEffectsAsync(),
                    LoadParetoAsync(),
                    LoadActualVsPredictedAsync(),
                    LoadResidualAsync(),
                    LoadAnovaAsync(),
                    LoadRegressionAsync(),
                    LoadSurfaceAsync()
                );

                StatusMessage = "图表加载完成";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载分析图表失败");
                StatusMessage = $"加载失败: {ex.Message}";
            }
            finally
            {
                IsLoading = false;
            }
        }

        // ══════════════ 各图表加载方法 ══════════════

        private async Task LoadMainEffectsAsync()
        {
            try
            {
                var json = await _analysisService.GetMainEffectsAsync(BatchId);
                var data = JsonConvert.DeserializeObject<Dictionary<string, List<LevelMean>>>(json);
                if (data == null || data.Count == 0) return;

                var model = new PlotModel { Title = "主效应图" };

                // ★ 修复: 检测是否存在类别因子（level 为字符串）
                // 混合因子时统一用索引做 X 轴，附加标签注释
                bool hasCategory = data.Values.Any(pts => pts.Any(p => !p.IsNumericLevel));

                if (hasCategory)
                {
                    // 有类别因子: 每个因子独立子图，X 轴用类别标签
                    // 简化处理: 所有因子共用一个图，X 轴用索引
                    model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "因子水平 (索引)" });
                }
                else
                {
                    model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "因子水平" });
                }
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "响应均值" });

                var colors = new[] { OxyColors.Blue, OxyColors.Red, OxyColors.Green, OxyColors.Orange, OxyColors.Purple };
                int ci = 0;

                foreach (var kv in data)
                {
                    var series = new LineSeries { Title = kv.Key, Color = colors[ci % colors.Length], MarkerType = MarkerType.Circle, MarkerSize = 5 };
                    for (int i = 0; i < kv.Value.Count; i++)
                    {
                        var pt = kv.Value[i];
                        // 连续因子用实际数值，类别因子用索引
                        double xVal = pt.IsNumericLevel ? pt.NumericLevel : i;
                        series.Points.Add(new DataPoint(xVal, pt.Mean));
                    }
                    model.Series.Add(series);
                    ci++;
                }

                MainEffectsPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "主效应图加载失败"); }
        }

        private async Task LoadParetoAsync()
        {
            try
            {
                var json = await _analysisService.GetParetoChartAsync(BatchId);
                var data = JsonConvert.DeserializeObject<List<ParetoItem>>(json);
                if (data == null || data.Count == 0) return;

                var sorted = data.OrderByDescending(d => d.Effect).ToList();
                var model = new PlotModel { Title = "Pareto 效应图", PlotMargins = new OxyThickness(10, 10, 40, 40) };

                var categoryAxis = new CategoryAxis { Position = AxisPosition.Bottom, GapWidth = 0.3 };
                foreach (var item in sorted) categoryAxis.Labels.Add(item.Term);

                var valueAxis = new LinearAxis { Position = AxisPosition.Left, Title = "效应大小", Minimum = 0, AbsoluteMinimum = 0 };

                model.Axes.Add(categoryAxis);
                model.Axes.Add(valueAxis);

                var series = new RectangleBarSeries { StrokeThickness = 0.5, StrokeColor = OxyColors.Gray };
                for (int i = 0; i < sorted.Count; i++)
                {
                    var item = sorted[i];
                    var color = item.Significant ? OxyColors.SteelBlue : OxyColors.LightGray;
                    series.Items.Add(new RectangleBarItem { X0 = i - 0.35, X1 = i + 0.35, Y0 = 0, Y1 = item.Effect, Color = color });
                }

                model.Series.Add(series);
                ParetoPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "Pareto 图加载失败"); }
        }

        private async Task LoadActualVsPredictedAsync()
        {
            try
            {
                var json = await _analysisService.GetActualVsPredictedAsync(BatchId);
                var data = JsonConvert.DeserializeObject<ActualVsPredicted>(json);
                if (data?.Actual == null) return;

                var model = new PlotModel { Title = $"实际值 vs 预测值 (R² = {data.RSquared:F4})" };
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "实际值" });
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "预测值" });

                var scatter = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 5, MarkerFill = OxyColors.SteelBlue };
                for (int i = 0; i < data.Actual.Count; i++)
                    scatter.Points.Add(new ScatterPoint(data.Actual[i], data.Predicted[i]));
                model.Series.Add(scatter);

                double min = Math.Min(data.Actual.Min(), data.Predicted.Min());
                double max = Math.Max(data.Actual.Max(), data.Predicted.Max());
                var diag = new LineSeries { Color = OxyColors.Red, LineStyle = LineStyle.Dash };
                diag.Points.Add(new DataPoint(min, min));
                diag.Points.Add(new DataPoint(max, max));
                model.Series.Add(diag);

                ActualVsPredictedPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "实际vs预测图加载失败"); }
        }

        private async Task LoadResidualAsync()
        {
            try
            {
                var json = await _analysisService.GetResidualAnalysisAsync(BatchId);
                var data = JsonConvert.DeserializeObject<ResidualData>(json);
                if (data?.Fitted == null) return;

                var model = new PlotModel { Title = "残差 vs 拟合值" };
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = "拟合值" });
                model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = "残差" });

                var scatter = new ScatterSeries { MarkerType = MarkerType.Circle, MarkerSize = 4, MarkerFill = OxyColors.DarkOrange };
                for (int i = 0; i < data.Fitted.Count; i++)
                    scatter.Points.Add(new ScatterPoint(data.Fitted[i], data.Residuals[i]));
                model.Series.Add(scatter);

                double min = data.Fitted.Min(), max = data.Fitted.Max();
                var zero = new LineSeries { Color = OxyColors.Gray, LineStyle = LineStyle.Dash };
                zero.Points.Add(new DataPoint(min, 0));
                zero.Points.Add(new DataPoint(max, 0));
                model.Series.Add(zero);

                ResidualPlot = model;
            }
            catch (Exception ex) { _logger.LogError(ex, "残差图加载失败"); }
        }

        private async Task LoadAnovaAsync()
        {
            try
            {
                var json = await _analysisService.GetAnovaTableAsync(BatchId);
                var data = JsonConvert.DeserializeObject<List<AnovaRowDto>>(json);
                if (data == null) return;

                var lines = new List<string> { "来源\t\tDF\tSS\t\tMS\t\tF\tP" };
                lines.Add(new string('-', 70));
                foreach (var row in data)
                    lines.Add($"{row.Source}\t\t{row.DF}\t{row.SS:F2}\t\t{row.MS:F2}\t\t{row.FValue?.ToString("F2") ?? "-"}\t{row.PValue?.ToString("F4") ?? "-"}");

                AnovaTableText = string.Join("\n", lines);
            }
            catch (Exception ex) { _logger.LogError(ex, "ANOVA 表加载失败"); }
        }

        private async Task LoadRegressionAsync()
        {
            try
            {
                var json = await _analysisService.GetRegressionSummaryAsync(BatchId);
                var data = JsonConvert.DeserializeObject<RegressionSummaryDto>(json);
                if (data == null) return;

                RegressionEquation = data.Equation ?? "";
                RSquaredText = $"R² = {data.RSquared:F4}, Adj-R² = {data.AdjRSquared:F4}";
            }
            catch (Exception ex) { _logger.LogError(ex, "回归摘要加载失败"); }
        }

        private async Task LoadSurfaceAsync()
        {
            if (string.IsNullOrEmpty(SurfaceFactor1) || string.IsNullOrEmpty(SurfaceFactor2) || SurfaceFactor1 == SurfaceFactor2)
                return;

            try
            {
                var imageBytes = await _analysisService.GetResponseSurfaceImageAsync(BatchId, SurfaceFactor1, SurfaceFactor2);
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
                    SurfaceImage = bmp;
                }
            }
            catch (Exception ex) { _logger.LogError(ex, "响应曲面图加载失败"); }
        }

        // ── JSON 映射 DTO ──

        private class LevelMean
        {
            /// <summary>
            /// ★ 修复: 从 double 改为 object，支持类别因子的字符串水平标签
            /// 连续因子: 数值 (如 120.0)，类别因子: 字符串 (如 "催化剂A")
            /// </summary>
            [JsonProperty("level")] public object Level { get; set; } = 0.0;
            [JsonProperty("mean")] public double Mean { get; set; }

            /// <summary>安全获取数值Level（类别因子返回索引值）</summary>
            public double NumericLevel
            {
                get
                {
                    if (Level is double d) return d;
                    if (Level is long l) return l;
                    if (Level is int i) return i;
                    if (double.TryParse(Level?.ToString(), out var parsed)) return parsed;
                    return 0.0;
                }
            }

            public bool IsNumericLevel => Level is double or long or int
                || double.TryParse(Level?.ToString(), out _);
        }

        private class ParetoItem
        {
            [JsonProperty("term")] public string Term { get; set; } = "";
            [JsonProperty("effect")] public double Effect { get; set; }
            [JsonProperty("significant")] public bool Significant { get; set; }
        }

        private class ActualVsPredicted
        {
            [JsonProperty("actual")] public List<double> Actual { get; set; } = new();
            [JsonProperty("predicted")] public List<double> Predicted { get; set; } = new();
            [JsonProperty("r_squared")] public double RSquared { get; set; }
        }

        private class ResidualData
        {
            [JsonProperty("fitted")] public List<double> Fitted { get; set; } = new();
            [JsonProperty("residuals")] public List<double> Residuals { get; set; } = new();
        }

        //  修复: 重命名避免和 Models.AnovaRow 冲突
        private class AnovaRowDto
        {
            [JsonProperty("source")] public string Source { get; set; } = "";
            [JsonProperty("df")] public int DF { get; set; }
            [JsonProperty("ss")] public double SS { get; set; }
            [JsonProperty("ms")] public double MS { get; set; }
            [JsonProperty("f_value")] public double? FValue { get; set; }
            [JsonProperty("p_value")] public double? PValue { get; set; }
        }

        private class RegressionSummaryDto
        {
            [JsonProperty("r_squared")] public double RSquared { get; set; }
            [JsonProperty("adj_r_squared")] public double AdjRSquared { get; set; }
            [JsonProperty("equation")] public string? Equation { get; set; }
        }
    }
}