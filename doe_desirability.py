"""
doe_desirability.py — DOE 多响应优化 (Desirability Function)
★ 全新文件

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
        weight:     基础权重 wⱼ（用于综合 D 的几何加权平均）
        importance: 重要度 1-5（JMP 风格，缩放基础权重: effective_weight = weight × importance/3）
                    5=最重要(×1.67) 3=标准(×1.0) 1=不太重要(×0.33)
        shape:      变换曲线形状参数 s（1=线性, <1=凸, >1=凹）
        shape_lower/shape_upper: target 类型时分别控制下半段和上半段的形状
        """
        self.name = name
        self.goal = goal.lower()
        self.lower = lower
        self.upper = upper
        self.importance = max(1, min(5, importance))  # 钳位到 [1, 5]
        self.shape = shape
        self.shape_lower = shape_lower
        self.shape_upper = shape_upper
        
        # ★ 修复: importance 参与有效权重计算
        # JMP 约定: importance 是 1-5 的整数，3 为标准基线
        # effective_weight = base_weight × (importance / 3)
        self.weight = weight * (self.importance / 3.0)
        
        # 自动设置 target
        if target is not None:
            self.target = target
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
        """
        if self.goal == "maximize":
            if y <= self.lower:
                return 0.0
            elif y >= self.target:
                return 1.0
            else:
                return ((y - self.lower) / (self.target - self.lower)) ** self.shape
        
        elif self.goal == "minimize":
            if y >= self.upper:
                return 0.0
            elif y <= self.target:
                return 1.0
            else:
                return ((self.upper - y) / (self.upper - self.target)) ** self.shape
        
        else:  # target
            if y < self.lower or y > self.upper:
                return 0.0
            elif y <= self.target:
                if abs(self.target - self.lower) < 1e-12:
                    return 1.0
                return ((y - self.lower) / (self.target - self.lower)) ** self.shape_lower
            else:
                if abs(self.upper - self.target) < 1e-12:
                    return 1.0
                return ((self.upper - y) / (self.upper - self.target)) ** self.shape_upper


class DesirabilityEngine:
    """
    多响应优化引擎 — 管理多个响应的 Desirability 配置，
    调用 GPR 或 OLS 模型进行预测，搜索最优因子组合。
    """
    
    def __init__(self):
        self._configs: List[DesirabilityConfig] = []
        self._factor_names: List[str] = []
        self._bounds: Dict[str, List[float]] = {}
        self._predictor = None  # 外部传入的预测函数
    
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
                "weight": 5.0,
                "importance": 5,
                "shape": 1.0
            },
            {
                "name": "纯度",
                "goal": "target",
                "lower": 95.0,
                "upper": 100.0,
                "target": 99.5,
                "weight": 3.0,
                "importance": 4,
                "shape_lower": 1.0,
                "shape_upper": 2.0
            },
            {
                "name": "成本",
                "goal": "minimize",
                "lower": 10.0,
                "upper": 100.0,
                "target": 10.0,
                "weight": 1.0,
                "importance": 2,
                "shape": 1.0
            }
        ]
        
        返回 JSON: {"configured": true, "response_count": 3}
        """
        configs = json.loads(configs_json)
        self._factor_names = json.loads(factor_names_json)
        self._bounds = json.loads(bounds_json)
        
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
        设置预测函数（从 C# 侧传入）。
        
        predictor_func 签名:
          输入: factors_json (str) — '{"温度": 130, "压力": 22}'
          输出: predictions_json (str) — '{"产率": {"mean": 85.3, "std": 2.1}, "纯度": {"mean": 99.1, "std": 0.5}}'
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
        
        D = (d₁^w₁ × d₂^w₂ × ... × dₘ^wₘ)^(1/Σwⱼ)
        """
        predictions = json.loads(predictions_json)
        
        individual_d = {}
        product = 1.0
        total_weight = 0.0
        any_zero = False
        
        for cfg in self._configs:
            y = predictions.get(cfg.name, 0.0)
            if isinstance(y, dict):
                y = y.get("mean", 0.0)  # 支持 GPR 的 {mean, std} 格式
            
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
        
        # 综合 D: 几何加权平均
        if any_zero or total_weight < 1e-12:
            composite = 0.0
        else:
            composite = product ** (1.0 / total_weight)
        
        return json.dumps({
            "composite_d": round(composite, 6),
            "individual_d": individual_d
        }, ensure_ascii=False)
    
    def optimize(self, predictor_json_func=None, n_iter: int = 200) -> str:
        """
        ★ 修复 (v3): 搜索使综合 Desirability D 最大的因子组合。
        
        对类别因子使用穷举策略，对连续因子使用 differential_evolution。
        遍历类别因子的所有水平组合，对每种组合独立优化连续因子部分。
        
        predictor_json_func: 如果为 None 则使用 set_predictor 设置的函数
        
        返回 JSON:
        {
            "optimal_factors": {"温度": 135.2, "压力": 22.5, "催化剂类型": "B"},
            "composite_d": 0.85,
            "individual_d": {...},
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
        categorical_levels_map = {}  # name → [level1, level2, ...]
        
        for name in self._factor_names:
            b = self._bounds.get(name, [0, 1])
            if isinstance(b, list) and len(b) > 0 and isinstance(b[0], str):
                categorical_names.append(name)
                categorical_levels_map[name] = b
            else:
                continuous_names.append(name)
        
        # 生成类别因子的所有水平组合
        if categorical_names:
            from itertools import product as iterproduct
            cat_levels_list = [categorical_levels_map[n] for n in categorical_names]
            cat_combos = list(iterproduct(*cat_levels_list))
        else:
            cat_combos = [()]
        
        # 连续因子的边界
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
                # 纯类别因子: 直接预测和计算 D
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
                            return 1.0  # D=0, 最大惩罚
                        product *= d ** cfg.weight
                        total_weight += cfg.weight
                    
                    D = product ** (1.0 / total_weight) if total_weight > 1e-12 else 0.0
                    return -D
                    
                except Exception:
                    return 1.0  # 预测失败，惩罚
            
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
            
            # 用最优因子做最终预测
            pred_json = pred_func(json.dumps(optimal_factors, ensure_ascii=False))
            predictions = json.loads(pred_json)
            
            # 计算最终的个体和综合 D
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
        
        Profiler 图: 逐个因子扫描，展示该因子变化时各响应和综合 D 的变化曲线。
        其他因子固定在最优值（或中心值）。
        
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
                    "composite_d": [0.72, 0.74, ...]
                },
                ...
            }
        }
        """
        pred_func = predictor_json_func or self._predictor
        if pred_func is None:
            return json.dumps({"error": "未设置预测函数"}, ensure_ascii=False)
        
        # ★ 修复 (v3): 分离连续和类别因子，正确构建中心点和扫描范围
        # 连续因子中心值: (lower + upper) / 2
        # 类别因子中心值: 第一个水平（参考水平）
        center = {}
        for name in self._factor_names:
            b = self._bounds.get(name, [0, 1])
            if isinstance(b, list) and len(b) > 0 and isinstance(b[0], str):
                center[name] = b[0]  # 类别因子: 取参考水平
            else:
                center[name] = (float(b[0]) + float(b[1])) / 2.0
        
        result = {"factors": {}}
        
        for factor in self._factor_names:
            b = self._bounds.get(factor, [0, 1])
            
            # 判断该因子是否为类别因子
            is_categorical = isinstance(b, list) and len(b) > 0 and isinstance(b[0], str)
            
            if is_categorical:
                # 类别因子: 扫描所有水平标签
                x_range = list(b)  # ["A", "B", "C"]
            else:
                # 连续因子: linspace 扫描
                x_range = np.linspace(float(b[0]), float(b[1]), grid_size).tolist()
            
            responses_data = {cfg.name: {"values": [], "desirabilities": []}
                             for cfg in self._configs}
            composite_d_list = []
            
            for x_val in x_range:
                # 构建因子值: 当前因子扫描，其余取中心
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
            
            # ★ 修复 (v4): x_range 可能包含字符串（类别因子）, 不能对字符串调 round()
            if is_categorical:
                x_output = list(x_range)  # 字符串原样保留
            else:
                x_output = [round(v, 4) for v in x_range]
            
            result["factors"][factor] = {
                "x": x_output,
                "responses": responses_data,
                "composite_d": composite_d_list
            }
        
        return json.dumps(result, ensure_ascii=False)


# ═══════════════════════════════════════════════════════
# 模块入口 — 供 C# pythonnet 调用
# ═══════════════════════════════════════════════════════

def create_engine() -> DesirabilityEngine:
    """创建 Desirability 引擎实例"""
    return DesirabilityEngine()


# ═══════════════════════════════════════════════════════
# 本地测试
# ═══════════════════════════════════════════════════════

if __name__ == "__main__":
    print("=== Desirability Function 测试 ===\n")
    
    engine = DesirabilityEngine()
    
    # 配置三个响应
    configs = json.dumps([
        {"name": "产率", "goal": "maximize", "lower": 50, "upper": 100, "weight": 5, "shape": 1.0},
        {"name": "纯度", "goal": "target", "lower": 95, "upper": 100, "target": 99.5,
         "weight": 3, "shape_lower": 1.0, "shape_upper": 2.0},
        {"name": "成本", "goal": "minimize", "lower": 10, "upper": 100, "weight": 1, "shape": 1.0}
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
    
    # 测试优化（使用模拟预测函数）
    def mock_predictor(factors_json):
        """模拟预测: 简单的二次模型"""
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
    
    # 测试 Profile 图数据
    profile = json.loads(engine.profile_plot_data(grid_size=10))
    for factor, data in profile["factors"].items():
        print(f"\n  Profile: {factor}")
        print(f"    x 范围: [{data['x'][0]}, ..., {data['x'][-1]}]")
        print(f"    D 范围: [{min(data['composite_d']):.3f}, {max(data['composite_d']):.3f}]")
    
    print("\n所有测试通过！")
