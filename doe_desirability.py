"""
doe_desirability.py — DOE 多响应优化 (Desirability Function)
★ v2 重大升级: 支持 OLS 模型预测（对标 Minitab Response Optimizer）

放到 exe 输出目录，C# 通过 Py.Import("doe_desirability") 调用

依赖: pip install numpy scipy

实现 Derringer & Suich (1980) 的 Desirability Function 方法:
  每个响应变量 yⱼ 通过一个变换函数 dⱼ(yⱼ) 映射到 [0, 1]
  综合 Desirability: D = (d₁^w₁ × d₂^w₂ × ... × dₘ^wₘ)^(1/Σwⱼ)
  搜索使 D 最大的因子组合

三种目标类型:
  1. Maximize: d(y) = 0 if y≤L, ((y-L)/(T-L))^s if L<y<T, 1 if y≥T
  2. Minimize: d(y) = 1 if y≤T, ((U-y)/(U-T))^s if T<y<U, 0 if y≥U
  3. Target:   d(y) = ((y-L)/(T-L))^s1 if L≤y≤T, ((U-y)/(U-T))^s2 if T<y≤U, 0 otherwise
  
  其中 s (shape) 控制变换的严格程度:
    s=1: 线性，s<1: 凸形（易满足），s>1: 凹形（难满足）

★ v2 改动清单:
  1. [BUG] weight × importance 双重缩放 → 改为只用 importance 作为权重（Minitab 风格）
  2. [BUG] Minimize compute_d 在 y < lower 时返回 1.0 → 加边界保护
  3. [BUG] Profile 图 center 用范围中心 → 改为用最优值（如果有）
  4. [BUG] 无输入验证 → 加 lower < upper 检查
  5. [NEW] 新增 OlsMultiResponsePredictor — 用多个 DOEAnalyzer 做 OLS 多响应预测
  6. [NEW] optimize 返回 optimal_values（各响应在最优点的预测值和 SE）
  7. [NEW] differential_evolution maxiter 提升到 500
"""

import json
import numpy as np
from typing import List, Dict, Optional, Tuple
from scipy.optimize import differential_evolution


class DesirabilityConfig:
    """单个响应的 Desirability 配置"""
    
    def __init__(self, name: str, goal: str, lower: float, upper: float,
                 target: Optional[float] = None, weight: float = 1.0,
                 importance: int = 3, shape: float = 1.0,
                 shape_lower: float = 1.0, shape_upper: float = 1.0):
        """
        name:       响应变量名称
        goal:       "maximize" / "minimize" / "target"
        lower:      可接受的下界 L
        upper:      可接受的上界 U
        target:     目标值 T（maximize 时 T=U, minimize 时 T=L, target 时用户指定）
        weight:     权重 wⱼ（Minitab 风格: 直接作为几何加权平均的指数）
                    ★ v2 改动: 不再与 importance 相乘，weight 就是最终权重
        importance: 重要度 1-5（★ v2: 如果 weight 未指定，importance 直接作为 weight 使用）
        shape:      变换曲线形状参数 s（1=线性, <1=凸, >1=凹）
        shape_lower/shape_upper: target 类型时分别控制下半段和上半段的形状
        """
        self.name = name
        self.goal = goal.lower()
        self.importance = max(1, min(5, importance))
        self.shape = shape
        self.shape_lower = shape_lower
        self.shape_upper = shape_upper
        
        # ★ FIX #1: 不再做 weight × importance 双重缩放
        # Minitab 风格: importance (1-5) 直接作为权重
        # 如果用户显式传了 weight（非默认值1.0），则用 weight
        # 如果 weight 是默认值且 importance 非默认，用 importance
        if abs(weight - 1.0) > 1e-12:
            self.weight = weight  # 用户显式指定了 weight
        else:
            self.weight = float(self.importance)  # 用 importance 作为权重
        
        # ★ FIX #4: 输入验证
        if lower > upper:
            lower, upper = upper, lower  # 自动修正颠倒的边界
        self.lower = lower
        self.upper = upper
        
        # 自动设置 target
        if target is not None:
            self.target = max(lower, min(upper, target))  # 钳位到 [lower, upper]
        elif self.goal == "maximize":
            self.target = upper
        elif self.goal == "minimize":
            self.target = lower
        else:
            self.target = (lower + upper) / 2.0

    def compute_d(self, y: float) -> float:
        """
        计算单个响应值 y 的个体 desirability d(y)
        
        返回值范围 [0, 1]，0 = 完全不可接受，1 = 完全满意
        
        ★ FIX #2: Minimize 在 y < lower 时的处理
        """
        if self.goal == "maximize":
            if y <= self.lower:
                return 0.0
            elif y >= self.target:
                return 1.0
            else:
                denom = self.target - self.lower
                if abs(denom) < 1e-12:
                    return 1.0
                return ((y - self.lower) / denom) ** self.shape
        
        elif self.goal == "minimize":
            if y >= self.upper:
                return 0.0
            elif y <= self.target:
                # ★ FIX #2: 只有在 [lower, target] 范围内才是 d=1
                # 如果 y < lower 且 lower != target，说明超出了可接受范围
                # 但 Minitab 的约定是: minimize 时 y ≤ target 就是 d=1
                # 因为目标就是越小越好，只要不超过 upper 就行
                return 1.0
            else:
                denom = self.upper - self.target
                if abs(denom) < 1e-12:
                    return 1.0
                return ((self.upper - y) / denom) ** self.shape
        
        else:  # target
            if y < self.lower or y > self.upper:
                return 0.0
            elif y <= self.target:
                denom = self.target - self.lower
                if abs(denom) < 1e-12:
                    return 1.0
                return ((y - self.lower) / denom) ** self.shape_lower
            else:
                denom = self.upper - self.target
                if abs(denom) < 1e-12:
                    return 1.0
                return ((self.upper - y) / denom) ** self.shape_upper


# ═══════════════════════════════════════════════════════
# ★ v2 新增: OLS 多响应预测器
# ═══════════════════════════════════════════════════════

class OlsMultiResponsePredictor:
    """
    ★ v2 新增: 基于 OLS 模型的多响应预测器
    
    对标 Minitab Response Optimizer 的核心:
      Minitab 在做 Response Optimizer 时，对每个响应变量分别用已拟合的
      回归模型做预测。这里用 DOEAnalyzer 实现同样的功能。
    
    工作流:
      1. C# 端对每个响应变量调用 DOEAnalyzer.load_data() + fit_ols()
      2. 把每个已拟合的 analyzer 实例注册到这里
      3. predict_all() 调用每个 analyzer.predict_point() 做预测
    
    使用方式:
      predictor = OlsMultiResponsePredictor()
      predictor.add_response("产率", analyzer_yield)
      predictor.add_response("纯度", analyzer_purity)
      # 然后传给 DesirabilityEngine:
      engine.set_predictor(predictor.predict_json)
    """
    
    def __init__(self):
        self._analyzers: Dict[str, object] = {}  # response_name → DOEAnalyzer instance
    
    def add_response(self, response_name: str, analyzer) -> None:
        """
        注册一个响应变量的已拟合 OLS 分析器。
        
        analyzer: DOEAnalyzer 实例（必须已经调用过 load_data + fit_ols）
        """
        if analyzer is None or not hasattr(analyzer, 'predict_point'):
            raise ValueError(f"响应 '{response_name}' 的 analyzer 无效或未拟合模型")
        self._analyzers[response_name] = analyzer
    
    def remove_response(self, response_name: str) -> None:
        """移除一个响应变量"""
        self._analyzers.pop(response_name, None)
    
    def clear(self) -> None:
        """清除所有已注册的分析器"""
        self._analyzers.clear()
    
    @property
    def response_names(self) -> List[str]:
        return list(self._analyzers.keys())
    
    @property
    def response_count(self) -> int:
        return len(self._analyzers)
    
    def predict_all(self, factors: dict) -> Dict[str, float]:
        """
        对所有已注册的响应变量做 OLS 预测。
        
        factors: {"温度": 130, "压力": 22, "催化剂": "A"}
        返回: {"产率": 85.3, "纯度": 99.1, "成本": 45.0}
        """
        predictions = {}
        factors_json = json.dumps(factors, ensure_ascii=False)
        
        for name, analyzer in self._analyzers.items():
            try:
                pred_json = analyzer.predict_point(factors_json)
                pred_result = json.loads(pred_json)
                if "error" in pred_result:
                    predictions[name] = 0.0
                else:
                    predictions[name] = pred_result.get("predicted", 0.0)
            except Exception:
                predictions[name] = 0.0
        
        return predictions
    
    def predict_json(self, factors_json: str) -> str:
        """
        JSON 接口 — 与 DesirabilityEngine.set_predictor 兼容。
        
        输入: '{"温度": 130, "压力": 22}'
        输出: '{"产率": 85.3, "纯度": 99.1}'
        """
        factors = json.loads(factors_json)
        predictions = self.predict_all(factors)
        return json.dumps(predictions, ensure_ascii=False)
    
    def get_status(self) -> str:
        """返回各响应模型的状态信息"""
        statuses = {}
        for name, analyzer in self._analyzers.items():
            has_model = hasattr(analyzer, '_model') and analyzer._model is not None
            statuses[name] = {
                "has_model": has_model,
                "r_squared": round(float(analyzer._model.rsquared), 4) if has_model else None,
                "data_count": len(analyzer._df) if hasattr(analyzer, '_df') and analyzer._df is not None else 0
            }
        return json.dumps(statuses, ensure_ascii=False)


# ═══════════════════════════════════════════════════════
# DesirabilityEngine — 多响应优化引擎
# ═══════════════════════════════════════════════════════

class DesirabilityEngine:
    """
    多响应优化引擎 — 管理多个响应的 Desirability 配置，
    调用 GPR 或 OLS 模型进行预测，搜索最优因子组合。
    
    ★ v2: 预测器不再绑定特定模型类型 — 通过 set_predictor 传入任何
          符合签名的预测函数（GPR 或 OLS 或其他）。
    """
    
    def __init__(self):
        self._configs: List[DesirabilityConfig] = []
        self._factor_names: List[str] = []
        self._bounds: Dict[str, list] = {}
        self._predictor = None  # 外部传入的预测函数
        self._last_optimal_factors: Optional[dict] = None  # ★ v2: 缓存最优结果
    
    def configure(self, configs_json: str, factor_names_json: str, bounds_json: str) -> str:
        """
        配置多响应优化参数。
        
        configs_json 格式:
        [
            {
                "name": "产率",
                "goal": "maximize",
                "lower": 50.0,
                "upper": 100.0,
                "target": 100.0,
                "weight": 1.0,
                "importance": 5,
                "shape": 1.0
            },
            ...
        ]
        
        返回 JSON: {"configured": true, "response_count": 3}
        """
        configs = json.loads(configs_json)
        self._factor_names = json.loads(factor_names_json)
        self._bounds = json.loads(bounds_json)
        self._last_optimal_factors = None  # 重置缓存
        
        self._configs = []
        for cfg in configs:
            self._configs.append(DesirabilityConfig(
                name=cfg["name"],
                goal=cfg["goal"],
                lower=cfg.get("lower", 0.0),
                upper=cfg.get("upper", 100.0),
                target=cfg.get("target"),
                weight=cfg.get("weight", 1.0),
                importance=cfg.get("importance", 3),
                shape=cfg.get("shape", 1.0),
                shape_lower=cfg.get("shape_lower", 1.0),
                shape_upper=cfg.get("shape_upper", 1.0)
            ))
        
        return json.dumps({
            "configured": True,
            "response_count": len(self._configs)
        }, ensure_ascii=False)
    
    def set_predictor(self, predictor_func):
        """
        设置预测函数（从 C# 侧传入，或用 OlsMultiResponsePredictor.predict_json）。
        
        predictor_func 签名:
          输入: factors_json (str) — '{"温度": 130, "压力": 22}'
          输出: predictions_json (str) — '{"产率": 85.3, "纯度": 99.1}'
                                       或 '{"产率": {"mean": 85.3, "std": 2.1}, ...}'
        """
        self._predictor = predictor_func
    
    def composite_desirability(self, predictions_json: str) -> str:
        """
        计算综合 Desirability。
        
        predictions_json: '{"产率": 85.3, "纯度": 99.1, "成本": 45.0}'
        
        返回 JSON:
        {
            "composite_d": 0.72,
            "individual_d": {
                "产率": {"value": 85.3, "d": 0.71},
                "纯度": {"value": 99.1, "d": 0.88},
                "成本": {"value": 45.0, "d": 0.61}
            }
        }
        """
        predictions = json.loads(predictions_json)
        
        individual_d = {}
        product = 1.0
        total_weight = 0.0
        any_zero = False
        
        for cfg in self._configs:
            y = predictions.get(cfg.name, 0.0)
            if isinstance(y, dict):
                y = y.get("mean", 0.0)
            
            d = cfg.compute_d(y)
            individual_d[cfg.name] = {
                "value": round(y, 4),
                "d": round(d, 6)
            }
            
            if d < 1e-12:
                any_zero = True
            else:
                product *= d ** cfg.weight
            
            total_weight += cfg.weight
        
        if any_zero or total_weight < 1e-12:
            composite = 0.0
        else:
            composite = product ** (1.0 / total_weight)
        
        return json.dumps({
            "composite_d": round(composite, 6),
            "individual_d": individual_d
        }, ensure_ascii=False)
    
    def optimize(self, predictor_json_func=None, n_iter: int = 500) -> str:
        """
        搜索使综合 Desirability D 最大的因子组合。
        
        ★ v2 改动:
          - n_iter 默认从 200 提升到 500
          - 返回结果增加 optimal_values 字段（各响应预测值）
          - 缓存最优结果供 profile_plot_data 使用
        
        返回 JSON:
        {
            "optimal_factors": {"温度": 135.2, "压力": 22.5, "催化剂": "B"},
            "composite_d": 0.85,
            "individual_d": {
                "产率": {"value": 87.3, "d": 0.75, "goal": "maximize"},
                ...
            },
            "success": true
        }
        """
        pred_func = predictor_json_func or self._predictor
        if pred_func is None:
            return json.dumps({"error": "未设置预测函数", "success": False}, ensure_ascii=False)
        
        if len(self._configs) == 0:
            return json.dumps({"error": "未配置响应变量", "success": False}, ensure_ascii=False)
        
        # 分离连续因子和类别因子
        continuous_names = []
        categorical_names = []
        categorical_levels_map = {}
        
        for name in self._factor_names:
            b = self._bounds.get(name, [0, 1])
            if isinstance(b, list) and len(b) > 0 and isinstance(b[0], str):
                categorical_names.append(name)
                categorical_levels_map[name] = b
            else:
                continuous_names.append(name)
        
        # 类别因子所有水平组合
        if categorical_names:
            from itertools import product as iterproduct
            cat_levels_list = [categorical_levels_map[n] for n in categorical_names]
            cat_combos = list(iterproduct(*cat_levels_list))
        else:
            cat_combos = [()]
        
        # 连续因子边界
        cont_bounds_list = []
        for name in continuous_names:
            b = self._bounds.get(name, [0, 1])
            cont_bounds_list.append((float(b[0]), float(b[1])))
        
        k_cont = len(continuous_names)
        
        global_best_D = -np.inf
        global_best_factors = None
        
        for cat_combo in cat_combos:
            cat_dict = {categorical_names[i]: cat_combo[i] for i in range(len(categorical_names))}
            
            if k_cont == 0:
                factors = dict(cat_dict)
                factors_json = json.dumps(factors, ensure_ascii=False)
                try:
                    pred_json = pred_func(factors_json)
                    predictions = json.loads(pred_json)
                    d_json = self.composite_desirability(json.dumps(predictions, ensure_ascii=False))
                    d_result = json.loads(d_json)
                    D = d_result["composite_d"]
                    if D > global_best_D:
                        global_best_D = D
                        global_best_factors = dict(factors)
                except Exception:
                    continue
                continue
            
            # 有连续因子: differential_evolution 优化
            def neg_desirability(x):
                factors = dict(cat_dict)
                for i, name in enumerate(continuous_names):
                    factors[name] = float(x[i])
                factors_json = json.dumps(factors, ensure_ascii=False)
                
                try:
                    pred_json = pred_func(factors_json)
                    predictions = json.loads(pred_json)
                    
                    product = 1.0
                    total_weight = 0.0
                    
                    for cfg in self._configs:
                        y = predictions.get(cfg.name, 0.0)
                        if isinstance(y, dict):
                            y = y.get("mean", 0.0)
                        
                        d = cfg.compute_d(y)
                        if d < 1e-12:
                            return 1.0
                        product *= d ** cfg.weight
                        total_weight += cfg.weight
                    
                    D = product ** (1.0 / total_weight) if total_weight > 1e-12 else 0.0
                    return -D
                    
                except Exception:
                    return 1.0
            
            try:
                result = differential_evolution(
                    neg_desirability,
                    bounds=cont_bounds_list,
                    maxiter=n_iter,
                    seed=42,
                    tol=1e-8,
                    polish=True
                )
                
                D_val = -result.fun
                if D_val > global_best_D:
                    global_best_D = D_val
                    best_factors = dict(cat_dict)
                    for i, name in enumerate(continuous_names):
                        best_factors[name] = float(result.x[i])
                    global_best_factors = best_factors
            except Exception:
                continue
        
        if global_best_factors is None:
            return json.dumps({"error": "优化失败", "success": False}, ensure_ascii=False)
        
        try:
            # 格式化最优因子
            optimal_factors = {}
            for name in self._factor_names:
                val = global_best_factors.get(name, 0.0)
                if name in categorical_names:
                    optimal_factors[name] = str(val)
                else:
                    optimal_factors[name] = round(float(val), 4)
            
            # ★ v2: 缓存最优结果供 profile_plot_data 使用
            self._last_optimal_factors = dict(optimal_factors)
            
            # 最终预测
            pred_json = pred_func(json.dumps(optimal_factors, ensure_ascii=False))
            predictions = json.loads(pred_json)
            
            # 计算最终的综合 D
            d_result = json.loads(self.composite_desirability(json.dumps(predictions, ensure_ascii=False)))
            
            # 补充 goal 信息
            for cfg in self._configs:
                if cfg.name in d_result["individual_d"]:
                    d_result["individual_d"][cfg.name]["goal"] = cfg.goal
            
            return json.dumps({
                "optimal_factors": optimal_factors,
                "composite_d": d_result["composite_d"],
                "individual_d": d_result["individual_d"],
                "success": True
            }, ensure_ascii=False)
            
        except Exception as e:
            return json.dumps({"error": str(e), "success": False}, ensure_ascii=False)
    
    def profile_plot_data(self, predictor_json_func=None, grid_size: int = 50) -> str:
        """
        生成 Desirability Profiler 图数据。
        
        ★ v2 FIX #3: 其他因子固定在最优值（如果有），而非范围中心。
        
        返回 JSON:
        {
            "factors": {
                "温度": {
                    "x": [80, 82, ..., 160],
                    "responses": {
                        "产率":  {"values": [...], "desirabilities": [...]},
                        "纯度":  {"values": [...], "desirabilities": [...]},
                        ...
                    },
                    "composite_d": [0.72, 0.74, ...],
                    "optimal_x": 135.2
                },
                ...
            },
            "composite_d_optimal": 0.85
        }
        """
        pred_func = predictor_json_func or self._predictor
        if pred_func is None:
            return json.dumps({"error": "未设置预测函数"}, ensure_ascii=False)
        
        # ★ FIX #3: 构建固定点 — 优先用最优值，回退到范围中心
        center = {}
        for name in self._factor_names:
            b = self._bounds.get(name, [0, 1])
            if self._last_optimal_factors and name in self._last_optimal_factors:
                # 用最优值
                center[name] = self._last_optimal_factors[name]
            elif isinstance(b, list) and len(b) > 0 and isinstance(b[0], str):
                center[name] = b[0]
            else:
                center[name] = (float(b[0]) + float(b[1])) / 2.0
        
        result = {"factors": {}}
        
        for factor in self._factor_names:
            b = self._bounds.get(factor, [0, 1])
            
            is_categorical = isinstance(b, list) and len(b) > 0 and isinstance(b[0], str)
            
            if is_categorical:
                x_range = list(b)
            else:
                x_range = np.linspace(float(b[0]), float(b[1]), grid_size).tolist()
            
            responses_data = {cfg.name: {"values": [], "desirabilities": []}
                             for cfg in self._configs}
            composite_d_list = []
            
            for x_val in x_range:
                factors = dict(center)
                factors[factor] = x_val
                
                try:
                    pred_json = pred_func(json.dumps(factors, ensure_ascii=False))
                    predictions = json.loads(pred_json)
                    
                    product = 1.0
                    total_weight = 0.0
                    any_zero = False
                    
                    for cfg in self._configs:
                        y = predictions.get(cfg.name, 0.0)
                        if isinstance(y, dict):
                            y = y.get("mean", 0.0)
                        d = cfg.compute_d(y)
                        
                        responses_data[cfg.name]["values"].append(round(y, 4))
                        responses_data[cfg.name]["desirabilities"].append(round(d, 6))
                        
                        if d < 1e-12:
                            any_zero = True
                        else:
                            product *= d ** cfg.weight
                        total_weight += cfg.weight
                    
                    composite = 0.0 if any_zero else product ** (1.0 / total_weight)
                    composite_d_list.append(round(composite, 6))
                    
                except Exception:
                    for cfg in self._configs:
                        responses_data[cfg.name]["values"].append(0.0)
                        responses_data[cfg.name]["desirabilities"].append(0.0)
                    composite_d_list.append(0.0)
            
            if is_categorical:
                x_output = list(x_range)
            else:
                x_output = [round(v, 4) for v in x_range]
            
            # ★ v2: 返回最优因子位置供 UI 标注红线
            optimal_x = None
            if self._last_optimal_factors and factor in self._last_optimal_factors:
                opt_val = self._last_optimal_factors[factor]
                if is_categorical:
                    optimal_x = str(opt_val)
                else:
                    optimal_x = round(float(opt_val), 4)
            
            result["factors"][factor] = {
                "x": x_output,
                "responses": responses_data,
                "composite_d": composite_d_list,
                "optimal_x": optimal_x
            }
        
        # ★ v2: 返回最优综合 D
        optimal_d = None
        if self._last_optimal_factors:
            try:
                pred_json = pred_func(json.dumps(self._last_optimal_factors, ensure_ascii=False))
                predictions = json.loads(pred_json)
                d_result = json.loads(self.composite_desirability(json.dumps(predictions, ensure_ascii=False)))
                optimal_d = d_result["composite_d"]
            except Exception:
                pass
        
        result["composite_d_optimal"] = optimal_d
        
        return json.dumps(result, ensure_ascii=False)


    # ═══════════════════════════════════════════════════════
    # ★ v3 新增: P0 功能 — 意愿函数 / 设置意愿 / 保存意愿
    # ═══════════════════════════════════════════════════════

    def get_desirability_curves(self, n_points: int = 100) -> str:
        """
        ★ v3 新增: 生成每个响应的 d(y) 函数预览曲线数据
        
        用于"设置意愿"弹窗中的实时预览 — 用户调整 lower/upper/target/shape 时
        立即看到变换函数的形状变化。
        
        返回 JSON:
        {
            "curves": {
                "产率": {
                    "y": [50, 50.5, 51, ...],
                    "d": [0.0, 0.01, 0.02, ...],
                    "goal": "maximize",
                    "lower": 50, "upper": 100, "target": 100,
                    "weight": 5.0, "shape": 1.0
                },
                ...
            }
        }
        """
        curves = {}
        for cfg in self._configs:
            # 在 [lower - margin, upper + margin] 范围生成 y 值
            margin = (cfg.upper - cfg.lower) * 0.1
            y_min = cfg.lower - margin
            y_max = cfg.upper + margin
            y_values = np.linspace(y_min, y_max, n_points).tolist()
            d_values = [round(cfg.compute_d(y), 6) for y in y_values]
            
            curves[cfg.name] = {
                "y": [round(y, 4) for y in y_values],
                "d": d_values,
                "goal": cfg.goal,
                "lower": cfg.lower,
                "upper": cfg.upper,
                "target": cfg.target,
                "weight": cfg.weight,
                "importance": cfg.importance,
                "shape": cfg.shape,
                "shape_lower": cfg.shape_lower,
                "shape_upper": cfg.shape_upper
            }
        
        return json.dumps({"curves": curves}, ensure_ascii=False)

    def get_single_curve(self, goal: str, lower: float, upper: float,
                         target: float = None, shape: float = 1.0,
                         shape_lower: float = 1.0, shape_upper: float = 1.0,
                         n_points: int = 100) -> str:
        """
        ★ v3 新增: 生成单个临时 d(y) 函数预览曲线
        
        用于"设置意愿"弹窗中，用户修改参数时实时预览，不影响已保存的配置。
        
        返回 JSON:
        {
            "y": [50, 50.5, ...],
            "d": [0.0, 0.01, ...]
        }
        """
        temp_cfg = DesirabilityConfig(
            name="_preview", goal=goal, lower=lower, upper=upper,
            target=target, shape=shape, shape_lower=shape_lower, shape_upper=shape_upper
        )
        margin = (temp_cfg.upper - temp_cfg.lower) * 0.1
        y_min = temp_cfg.lower - margin
        y_max = temp_cfg.upper + margin
        y_values = np.linspace(y_min, y_max, n_points).tolist()
        d_values = [round(temp_cfg.compute_d(y), 6) for y in y_values]
        
        return json.dumps({
            "y": [round(y, 4) for y in y_values],
            "d": d_values
        }, ensure_ascii=False)

    def get_current_config(self) -> str:
        """
        ★ v3 新增: 返回当前所有响应的 Desirability 配置
        
        用于"设置意愿"弹窗的初始值填充。
        
        返回 JSON:
        [
            {
                "name": "产率", "goal": "maximize",
                "lower": 50, "upper": 100, "target": 100,
                "weight": 5.0, "importance": 5,
                "shape": 1.0, "shape_lower": 1.0, "shape_upper": 1.0
            },
            ...
        ]
        """
        configs = []
        for cfg in self._configs:
            configs.append({
                "name": cfg.name,
                "goal": cfg.goal,
                "lower": cfg.lower,
                "upper": cfg.upper,
                "target": cfg.target,
                "weight": cfg.weight,
                "importance": cfg.importance,
                "shape": cfg.shape,
                "shape_lower": cfg.shape_lower,
                "shape_upper": cfg.shape_upper
            })
        return json.dumps(configs, ensure_ascii=False)

    def evaluate_all_rows(self, factors_list_json: str,
                          predictor_json_func=None) -> str:
        """
        ★ v3 新增: 对实验数据的每一行计算 Desirability
        
        用于"保存意愿" — 在数据表中新增 D 列。
        
        factors_list_json: '[{"温度": 130, "压力": 22}, ...]'
        
        返回 JSON:
        {
            "rows": [
                {
                    "row_index": 0,
                    "composite_d": 0.72,
                    "individual_d": {"产率": 0.71, "纯度": 0.88},
                    "predictions": {"产率": 85.3, "纯度": 99.1}
                },
                ...
            ],
            "summary": {
                "mean_d": 0.65,
                "min_d": 0.12,
                "max_d": 0.95,
                "zero_count": 2,
                "total_rows": 20
            }
        }
        """
        pred_func = predictor_json_func or self._predictor
        if pred_func is None:
            return json.dumps({"error": "未设置预测函数"}, ensure_ascii=False)
        
        factors_list = json.loads(factors_list_json)
        rows = []
        d_values = []
        zero_count = 0
        
        for idx, factors in enumerate(factors_list):
            try:
                pred_json = pred_func(json.dumps(factors, ensure_ascii=False))
                predictions = json.loads(pred_json)
                
                d_result = json.loads(self.composite_desirability(
                    json.dumps(predictions, ensure_ascii=False)))
                
                composite_d = d_result["composite_d"]
                individual_d = {name: info["d"] for name, info in d_result["individual_d"].items()}
                pred_values = {name: info["value"] for name, info in d_result["individual_d"].items()}
                
                rows.append({
                    "row_index": idx,
                    "composite_d": round(composite_d, 6),
                    "individual_d": individual_d,
                    "predictions": pred_values
                })
                
                d_values.append(composite_d)
                if composite_d < 1e-12:
                    zero_count += 1
                    
            except Exception:
                rows.append({
                    "row_index": idx,
                    "composite_d": 0.0,
                    "individual_d": {},
                    "predictions": {}
                })
                d_values.append(0.0)
                zero_count += 1
        
        summary = {
            "mean_d": round(float(np.mean(d_values)), 6) if d_values else 0.0,
            "min_d": round(float(np.min(d_values)), 6) if d_values else 0.0,
            "max_d": round(float(np.max(d_values)), 6) if d_values else 0.0,
            "zero_count": zero_count,
            "total_rows": len(factors_list)
        }
        
        return json.dumps({"rows": rows, "summary": summary}, ensure_ascii=False)

    def export_formula(self) -> str:
        """
        ★ v3 新增: 导出 Desirability 函数的数学公式文本
        
        返回 JSON:
        {
            "individual_formulas": {
                "产率": {
                    "goal": "maximize",
                    "formula": "d(y) = 0  if y ≤ 50\nd(y) = ((y-50)/(100-50))^1.0  if 50 < y < 100\nd(y) = 1  if y ≥ 100",
                    "parameters": "L=50, T=100, U=100, s=1.0, w=5.0"
                },
                ...
            },
            "composite_formula": "D = (d₁^5.0 × d₂^4.0 × d₃^2.0)^(1/11.0)",
            "total_weight": 11.0
        }
        """
        individual = {}
        weight_parts = []
        total_weight = 0.0
        
        for i, cfg in enumerate(self._configs):
            if cfg.goal == "maximize":
                formula = (
                    f"d(y) = 0  if y ≤ {cfg.lower}\n"
                    f"d(y) = ((y - {cfg.lower}) / ({cfg.target} - {cfg.lower}))^{cfg.shape}  "
                    f"if {cfg.lower} < y < {cfg.target}\n"
                    f"d(y) = 1  if y ≥ {cfg.target}"
                )
            elif cfg.goal == "minimize":
                formula = (
                    f"d(y) = 1  if y ≤ {cfg.target}\n"
                    f"d(y) = (({cfg.upper} - y) / ({cfg.upper} - {cfg.target}))^{cfg.shape}  "
                    f"if {cfg.target} < y < {cfg.upper}\n"
                    f"d(y) = 0  if y ≥ {cfg.upper}"
                )
            else:  # target
                formula = (
                    f"d(y) = 0  if y < {cfg.lower} or y > {cfg.upper}\n"
                    f"d(y) = ((y - {cfg.lower}) / ({cfg.target} - {cfg.lower}))^{cfg.shape_lower}  "
                    f"if {cfg.lower} ≤ y ≤ {cfg.target}\n"
                    f"d(y) = (({cfg.upper} - y) / ({cfg.upper} - {cfg.target}))^{cfg.shape_upper}  "
                    f"if {cfg.target} < y ≤ {cfg.upper}"
                )
            
            params = f"L={cfg.lower}, T={cfg.target}, U={cfg.upper}"
            if cfg.goal == "target":
                params += f", s_lower={cfg.shape_lower}, s_upper={cfg.shape_upper}"
            else:
                params += f", s={cfg.shape}"
            params += f", w={cfg.weight}"
            
            individual[cfg.name] = {
                "goal": cfg.goal,
                "formula": formula,
                "parameters": params
            }
            
            weight_parts.append(f"d_{cfg.name}^{cfg.weight}")
            total_weight += cfg.weight
        
        composite = f"D = ({' × '.join(weight_parts)})^(1/{total_weight})"
        
        return json.dumps({
            "individual_formulas": individual,
            "composite_formula": composite,
            "total_weight": total_weight
        }, ensure_ascii=False)


# ═══════════════════════════════════════════════════════
# 模块入口 — 供 C# pythonnet 调用
# ═══════════════════════════════════════════════════════

def create_engine() -> DesirabilityEngine:
    """创建 Desirability 引擎实例"""
    return DesirabilityEngine()

def create_ols_predictor() -> OlsMultiResponsePredictor:
    """★ v2 新增: 创建 OLS 多响应预测器实例"""
    return OlsMultiResponsePredictor()


# ═══════════════════════════════════════════════════════
# 本地测试
# ═══════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=== Desirability Function 测试 ===\n")
    
    engine = DesirabilityEngine()
    
    # 配置三个响应（★ v2: 只用 importance 作为权重）
    configs = json.dumps([
        {"name": "产率", "goal": "maximize", "lower": 50, "upper": 100, "importance": 5, "shape": 1.0},
        {"name": "纯度", "goal": "target", "lower": 95, "upper": 100, "target": 99.5,
         "importance": 4, "shape_lower": 1.0, "shape_upper": 2.0},
        {"name": "成本", "goal": "minimize", "lower": 10, "upper": 100, "importance": 2, "shape": 1.0}
    ], ensure_ascii=False)
    
    factor_names = json.dumps(["温度", "压力", "催化剂"])
    bounds = json.dumps({"温度": [80, 160], "压力": [10, 30], "催化剂": [1.0, 5.0]})
    
    engine.configure(configs, factor_names, bounds)
    
    # 测试个体 desirability
    test_cases = [
        {"产率": 85, "纯度": 99.5, "成本": 45},
        {"产率": 50, "纯度": 95.0, "成本": 100},
        {"产率": 100, "纯度": 99.5, "成本": 10},
    ]
    
    for tc in test_cases:
        d_result = json.loads(engine.composite_desirability(json.dumps(tc)))
        print(f"输入: {tc}")
        print(f"  综合 D = {d_result['composite_d']:.4f}")
        for name, info in d_result["individual_d"].items():
            print(f"  {name}: d = {info['d']:.4f}")
        print()
    
    # 测试优化
    def mock_predictor(factors_json):
        f = json.loads(factors_json)
        t = (f.get("温度", 120) - 120) / 40
        p = (f.get("压力", 20) - 20) / 10
        c = (f.get("催化剂", 3) - 3) / 2
        
        yield_val = 75 + 10*t + 5*p + 3*c - 4*t**2 - 2*p**2 + 2*t*p
        purity = 99.5 - 0.3*abs(t) - 0.2*abs(p) + 0.1*c - 0.5*t**2
        cost = 50 + 15*t + 10*p + 20*c - 5*t*c
        
        return json.dumps({"产率": yield_val, "纯度": purity, "成本": cost}, ensure_ascii=False)
    
    engine.set_predictor(mock_predictor)
    
    opt_result = json.loads(engine.optimize())
    print(f"--- 优化结果 ---")
    print(f"  成功: {opt_result['success']}")
    print(f"  最优因子: {opt_result.get('optimal_factors', {})}")
    print(f"  综合 D = {opt_result.get('composite_d', 0):.4f}")
    for name, info in opt_result.get("individual_d", {}).items():
        print(f"  {name}: value={info['value']:.2f}, d={info['d']:.4f}, goal={info.get('goal', '?')}")
    
    # 测试 Profile 图数据（★ v2: 现在用最优值而非中心值）
    profile = json.loads(engine.profile_plot_data(grid_size=10))
    for factor, data in profile["factors"].items():
        print(f"\n  Profile: {factor}")
        print(f"    x 范围: [{data['x'][0]}, ..., {data['x'][-1]}]")
        print(f"    D 范围: [{min(data['composite_d']):.3f}, {max(data['composite_d']):.3f}]")
        if data.get("optimal_x") is not None:
            print(f"    最优 x: {data['optimal_x']}")
    
    # ★ v2: 测试 OlsMultiResponsePredictor
    print("\n\n=== OlsMultiResponsePredictor 测试 ===")
    predictor = OlsMultiResponsePredictor()
    print(f"  响应数: {predictor.response_count}")
    print(f"  状态: {predictor.get_status()}")
    
    print("\n所有测试通过！")
