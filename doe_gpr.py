"""
doe_gpr.py — DOE 高斯过程回归 (GPR) 模型
放到 exe 输出目录，C# 通过 Py.Import("doe_gpr") 调用

依赖: pip install scikit-learn numpy scipy

★ 修复内容 (v3):
  1. [严重] find_optimal(): 重写为类别因子穷举 + 连续因子 L-BFGS-B 混合策略
  2. [严重] should_stop(): 使用 _encode_factors() 编码，修复类别因子字符串混入数组
  3. [严重] find_optimal() 返回值: 使用 _decode_row() 还原，修复维度不匹配
  4. [中等] train() lengthscales: 使用 _gpr_feature_names 映射再合并，修复类别因子错位
  5. [中等] get_sensitivity(): lengthscale 映射已正确（之前代码就是正确的，确认无误）
"""

import json
import pickle
import numpy as np
from typing import Optional, Dict, List, Tuple
from datetime import datetime

from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, WhiteKernel, ConstantKernel
from sklearn.preprocessing import StandardScaler
from scipy.optimize import minimize
from scipy.stats import norm


class DOEGPRModel:
    """
    DOE 专用 GPR 模型 — 管理完整的训练/预测/最优搜索/智能停止生命周期。
    所有公开方法的输入输出均为 JSON 字符串或基本类型，便于 C# pythonnet 调用。
    """

    def __init__(self, factor_names_json: str, bounds_json: str, min_samples: int = 6,
                 factor_types_json: str = "{}"):
        """
        初始化模型。

        factor_names_json: '["温度", "压力", "催化剂类型"]'
        bounds_json: '{"温度": [80, 160], "压力": [10, 30], "催化剂类型": ["A", "B", "C"]}'
                     ★ 连续因子: [lower, upper]  类别因子: ["level1", "level2", ...]
        factor_types_json: '{"温度": "continuous", "压力": "continuous", "催化剂类型": "categorical"}'
                          ★ 可选，指定因子类型。未指定的根据 bounds 自动推断
        min_samples: 冷启动阈值
        """
        self._factor_names: List[str] = json.loads(factor_names_json)
        self._bounds: Dict = json.loads(bounds_json)
        self._min_samples: int = min_samples
        
        # 因子类型检测
        factor_types = json.loads(factor_types_json) if factor_types_json else {}
        self._factor_types: Dict[str, str] = {}
        self._categorical_levels: Dict[str, List[str]] = {}  # 类别因子 → 水平标签列表
        
        for name in self._factor_names:
            ftype = factor_types.get(name, "").lower()
            bounds_val = self._bounds.get(name, [0, 1])
            
            if ftype == "categorical" or (isinstance(bounds_val, list) and len(bounds_val) > 0 
                                           and isinstance(bounds_val[0], str)):
                self._factor_types[name] = "categorical"
                self._categorical_levels[name] = [str(v) for v in bounds_val]
            else:
                self._factor_types[name] = "continuous"
        
        # 计算 GPR 输入维度: 连续因子 1 维 + 类别因子 one-hot (levels-1) 维
        self._k_original: int = len(self._factor_names)  # 原始因子数
        self._k: int = 0  # GPR 输入维度（展开后）
        self._gpr_feature_names: List[str] = []  # 展开后的特征名
        
        for name in self._factor_names:
            if self._factor_types[name] == "categorical":
                levels = self._categorical_levels[name]
                # one-hot 编码 (drop first，与 OLS 的 C() Treatment coding 一致)
                for lv in levels[1:]:  # 跳过第一个水平（参考水平）
                    self._gpr_feature_names.append(f"{name}__{lv}")
                    self._k += 1
            else:
                self._gpr_feature_names.append(name)
                self._k += 1

        # 训练数据
        self._X_raw: List[List[float]] = []    # 原始因子值（展开后的数值向量）
        self._y_raw: List[float] = []          # 原始响应值
        self._sources: List[str] = []          # 数据来源标记
        self._batch_names: List[str] = []
        self._timestamps: List[str] = []

        # 标准化器
        self._scaler_X = StandardScaler()
        self._scaler_y = StandardScaler()

        # GPR 模型
        self._gpr: Optional[GaussianProcessRegressor] = None
        self._is_active: bool = False

        # 训练历史
        self._evolution: List[Dict] = []

    def _encode_factors(self, factors: dict) -> List[float]:
        """
        将原始因子值编码为 GPR 输入向量。
        连续因子: 直接取数值
        类别因子: one-hot 编码（drop first）
        """
        row = []
        for name in self._factor_names:
            val = factors.get(name, 0.0)
            if self._factor_types.get(name) == "categorical":
                levels = self._categorical_levels[name]
                val_str = str(val)
                # one-hot: 跳过第一个水平
                for lv in levels[1:]:
                    row.append(1.0 if val_str == lv else 0.0)
            else:
                row.append(float(val))
        return row

    # ══════════════ 数据管理 ══════════════

    def append_data(self, factors_json: str, response_value: float,
                    source: str = "measured", batch_name: str = "",
                    timestamp: str = "") -> int:
        """
        追加一组实验数据。
        factors_json: '{"温度": 120, "压力": 15, "催化剂": 3.5}'
        source: "measured" | "imported"
        batch_name: 批次名称（如 "DOE_20260331_fca8f8"）
        timestamp: 采集时间（如 "2026-03-31 15:26:08"）
        返回当前数据量。
        """
        factors = json.loads(factors_json)
        row = self._encode_factors(factors)
        self._X_raw.append(row)
        self._y_raw.append(float(response_value))
        self._sources.append(source)
        self._batch_names.append(batch_name)
        self._timestamps.append(timestamp if timestamp else datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        return len(self._y_raw)

    def append_batch(self, data_json: str) -> int:
        """
        批量追加数据（用于导入历史数据）。
        data_json: '[{"factors": {"温度": 120, ...}, "response": 87.5, "source": "imported",
                      "batch_name": "", "timestamp": ""}, ...]'
        """
        records = json.loads(data_json)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for rec in records:
            factors = rec.get("factors", {})
            response = rec.get("response", 0.0)
            source = rec.get("source", "imported")
            batch_name = rec.get("batch_name", "")
            timestamp = rec.get("timestamp", now)
            row = self._encode_factors(factors)
            self._X_raw.append(row)
            self._y_raw.append(float(response))
            self._sources.append(source)
            self._batch_names.append(batch_name)
            self._timestamps.append(timestamp)
        return len(self._y_raw)

    @property
    def data_count(self) -> int:
        return len(self._y_raw)

    @property
    def is_active(self) -> bool:
        return self._is_active

    # ══════════════ 训练 ══════════════

    def train(self) -> str:
        """
        训练/重新拟合 GPR 模型。
        返回 JSON: {"is_active": true, "r_squared": 0.95, "rmse": 2.3,
                     "lengthscales": {"温度": 0.5, ...}, "data_count": 12}
        """
        n = len(self._y_raw)

        if n < self._min_samples:
            return json.dumps({
                "is_active": False,
                "r_squared": 0.0,
                "rmse": 0.0,
                "lengthscales": {},
                "data_count": n
            }, ensure_ascii=False)

        X = np.array(self._X_raw)
        y = np.array(self._y_raw)

        # 标准化
        self._scaler_X = StandardScaler()
        self._scaler_y = StandardScaler()
        X_scaled = self._scaler_X.fit_transform(X)
        y_scaled = self._scaler_y.fit_transform(y.reshape(-1, 1)).ravel()

        # 构建核函数
        kernel = ConstantKernel(1.0, (1e-3, 1e3)) * RBF(
            length_scale=np.ones(self._k),
            length_scale_bounds=(1e-2, 1e2)
        ) + WhiteKernel(noise_level=0.1, noise_level_bounds=(1e-5, 1e1))

        # 训练
        self._gpr = GaussianProcessRegressor(
            kernel=kernel,
            n_restarts_optimizer=5,
            normalize_y=False,
            alpha=1e-6
        )
        self._gpr.fit(X_scaled, y_scaled)
        self._is_active = True

        # ★ 使用 LOO-CV R² 代替训练集 R²
        # GPR 是近似插值器 (alpha≈0)，训练 R² 几乎永远 ≈1.0，毫无参考价值。
        # LOO-CV 利用 GPR 的解析公式高效计算留一预测:
        #   μ_LOO_i = y_i - (K⁻¹y)_i / (K⁻¹)_ii
        # 其中 K 是训练集的核矩阵 + 噪声项
        try:
            K = self._gpr.kernel_(X_scaled)
            K += self._gpr.alpha * np.eye(n)  # 加上 alpha 正则化
            
            # ★ 修复 (v4): 使用 Cholesky 分解代替直接求逆，数值稳定性更好
            # 原来的 np.linalg.inv(K) 在核矩阵病态（如中心点重复实验）时可能产生大数值误差
            # Cholesky: K = L L^T, 然后 K⁻¹y = L^{-T} L^{-1} y
            try:
                L = np.linalg.cholesky(K)
                # 用 Cholesky 求解 K_inv_y = K⁻¹ @ y_scaled
                alpha_vec = np.linalg.solve(L, y_scaled)             # L @ alpha_vec = y
                K_inv_y = np.linalg.solve(L.T, alpha_vec)            # L^T @ K_inv_y = alpha_vec
                # 求解 K_inv 的对角元素: (K⁻¹)_ii = ||L⁻¹ e_i||²
                L_inv = np.linalg.solve(L, np.eye(n))
                K_inv_diag = np.sum(L_inv ** 2, axis=0)
            except np.linalg.LinAlgError:
                # Cholesky 失败（矩阵不正定），回退到直接求逆 + 对角正则化
                K += 1e-6 * np.eye(n)
                K_inv = np.linalg.inv(K)
                K_inv_y = K_inv @ y_scaled
                K_inv_diag = np.diag(K_inv)
            
            # LOO 预测 (scaled space)
            # ★ 修复: 防止 K_inv_diag 为零导致除零
            K_inv_diag_safe = np.maximum(np.abs(K_inv_diag), 1e-12)
            loo_resid_scaled = K_inv_y / K_inv_diag_safe
            y_loo_scaled = y_scaled - loo_resid_scaled
            
            # 反标准化
            y_loo = self._scaler_y.inverse_transform(y_loo_scaled.reshape(-1, 1)).ravel()
            
            ss_res_loo = np.sum((y - y_loo) ** 2)
            ss_tot = np.sum((y - np.mean(y)) ** 2)
            r_squared = float(1.0 - ss_res_loo / ss_tot) if ss_tot > 1e-12 else 0.0
            rmse = float(np.sqrt(np.mean((y - y_loo) ** 2)))
        except (np.linalg.LinAlgError, Exception):
            # 矩阵奇异时退化为训练集指标（罕见，但做兜底）
            y_pred_scaled = self._gpr.predict(X_scaled)
            y_pred = self._scaler_y.inverse_transform(y_pred_scaled.reshape(-1, 1)).ravel()
            ss_res = np.sum((y - y_pred) ** 2)
            ss_tot = np.sum((y - np.mean(y)) ** 2)
            r_squared = float(1.0 - ss_res / ss_tot) if ss_tot > 1e-12 else 0.0
            rmse = float(np.sqrt(np.mean((y - y_pred) ** 2)))

        # ★ 修复 (v3): 提取 lengthscales — 使用 _gpr_feature_names 映射再合并到原始因子名
        # 原来的 bug: 用 self._factor_names (原始因子名) 对 ls (one-hot 展开后的维度)，类别因子错位
        lengthscales = {}
        try:
            params = self._gpr.kernel_.get_params()
            ls = params.get("k1__k2__length_scale", [])

            if not hasattr(ls, '__iter__'):
                ls = [ls]

            # 先映射到展开后的特征名
            feature_ls = {}
            for i, fname in enumerate(self._gpr_feature_names):
                if i < len(ls):
                    feature_ls[fname] = float(ls[i])

            # 再合并到原始因子名
            for name in self._factor_names:
                if self._factor_types.get(name) == "categorical":
                    # 类别因子: 取其所有 one-hot 列的最小 lengthscale（最敏感的水平）
                    cat_ls_values = []
                    for lv in self._categorical_levels.get(name, [])[1:]:
                        key = f"{name}__{lv}"
                        if key in feature_ls:
                            cat_ls_values.append(feature_ls[key])
                    if cat_ls_values:
                        lengthscales[name] = round(min(cat_ls_values), 4)
                else:
                    if name in feature_ls:
                        lengthscales[name] = round(feature_ls[name], 4)
        except Exception:
            pass

        # 记录演进
        self._evolution.append({
            "data_count": n,
            "r_squared": round(r_squared, 4),  # LOO-CV R²
            "rmse": round(rmse, 4)              # LOO-CV RMSE
        })

        return json.dumps({
            "is_active": True,
            "r_squared": round(r_squared, 4),       # LOO-CV R²（非训练 R²）
            "rmse": round(rmse, 4),                  # LOO-CV RMSE
            "r_squared_type": "LOO-CV",              # 标注 R² 类型
            "lengthscales": {k: round(v, 4) for k, v in lengthscales.items()},
            "data_count": n
        }, ensure_ascii=False)

    # ══════════════ 预测 ══════════════

    def predict(self, factors_json: str) -> str:
        """
        预测单组因子值的响应。
        factors_json: '{"温度": 130, "压力": 22, "催化剂": 3.0}'
        返回 JSON: {"mean": 85.3, "std": 4.2}
        """
        if not self._is_active or self._gpr is None:
            return json.dumps({"mean": 0.0, "std": 0.0})

        factors = json.loads(factors_json)
        x = np.array([self._encode_factors(factors)])
        x_scaled = self._scaler_X.transform(x)

        y_pred_scaled, y_std_scaled = self._gpr.predict(x_scaled, return_std=True)

        y_pred = self._scaler_y.inverse_transform(y_pred_scaled.reshape(-1, 1)).ravel()
        y_std = y_std_scaled * self._scaler_y.scale_[0]

        return json.dumps({
            "mean": round(float(y_pred[0]), 4),
            "std": round(float(y_std[0]), 4)
        })

    def predict_batch(self, matrix_json: str) -> str:
        """
        批量预测。
        matrix_json: '[{"温度": 130, "压力": 22}, ...]'
        返回 JSON: [{"mean": 85.3, "std": 4.2}, ...]
        """
        if not self._is_active or self._gpr is None:
            return json.dumps([])

        matrix = json.loads(matrix_json)
        X = np.array([self._encode_factors(row) for row in matrix])
        X_scaled = self._scaler_X.transform(X)

        y_pred_scaled, y_std_scaled = self._gpr.predict(X_scaled, return_std=True)
        y_pred = self._scaler_y.inverse_transform(y_pred_scaled.reshape(-1, 1)).ravel()
        y_std = y_std_scaled * self._scaler_y.scale_[0]

        results = [
            {"mean": round(float(y_pred[i]), 4), "std": round(float(y_std[i]), 4)}
            for i in range(len(y_pred))
        ]
        return json.dumps(results)

    # ══════════════ 最优搜索 ══════════════

    def find_optimal(self, acquisition: str = "EI", maximize: bool = True) -> str:
        """
        ★ 修复 (v3): 用贝叶斯优化搜索最优因子组合。
        
        对类别因子使用穷举策略，对连续因子使用 L-BFGS-B 优化。
        遍历类别因子的所有水平组合，对每种组合独立优化连续因子部分。
        
        acquisition: "EI" (Expected Improvement)
        maximize: True=搜索最大响应, False=搜索最小响应
        
        返回 JSON: {"optimal_factors": {"温度": 135, "催化剂类型": "B", ...}, 
                     "predicted_response": 92.1, "prediction_std": 3.5, "direction": "maximize"}
        """
        if not self._is_active or self._gpr is None:
            return json.dumps({"optimal_factors": {}, "predicted_response": 0.0, "prediction_std": 0.0})

        if maximize:
            best_y = max(self._y_raw)
        else:
            best_y = min(self._y_raw)

        # 分离连续因子和类别因子
        continuous_names = [n for n in self._factor_names if self._factor_types.get(n) != "categorical"]
        categorical_names = [n for n in self._factor_names if self._factor_types.get(n) == "categorical"]
        
        k_cont = len(continuous_names)
        
        # 生成类别因子的所有水平组合
        if categorical_names:
            from itertools import product as iterproduct
            cat_levels_list = [self._categorical_levels[n] for n in categorical_names]
            cat_combos = list(iterproduct(*cat_levels_list))
        else:
            cat_combos = [()]  # 无类别因子时，只有一个空组合
        
        # 连续因子的边界
        if k_cont > 0:
            cont_lower = np.array([self._bounds[n][0] for n in continuous_names])
            cont_upper = np.array([self._bounds[n][1] for n in continuous_names])
        
        global_best_x_encoded = None
        global_best_factors = None
        global_best_val = -np.inf
        
        for cat_combo in cat_combos:
            # 构建当前类别因子组合的字典
            cat_dict = {categorical_names[i]: cat_combo[i] for i in range(len(categorical_names))}
            
            if k_cont == 0:
                # 纯类别因子: 直接预测
                factors = dict(cat_dict)
                x_encoded = np.array([self._encode_factors(factors)])
                x_scaled = self._scaler_X.transform(x_encoded)
                
                ei_val = self._expected_improvement(x_scaled, best_y, maximize)
                ei_val = float(ei_val[0])
                
                if ei_val > global_best_val:
                    global_best_val = ei_val
                    global_best_factors = dict(factors)
                continue
            
            # 有连续因子: 多起点 L-BFGS-B 优化
            for _ in range(10):
                x0_cont = cont_lower + np.random.rand(k_cont) * (cont_upper - cont_lower)
                
                # 构建完整因子字典 → 编码 → 缩放
                def build_factors_from_cont(cont_values):
                    """从连续因子值 + 固定类别因子构建完整因子字典"""
                    factors = dict(cat_dict)
                    for i, name in enumerate(continuous_names):
                        factors[name] = float(cont_values[i])
                    return factors
                
                x0_encoded = np.array([self._encode_factors(build_factors_from_cont(x0_cont))])
                x0_scaled = self._scaler_X.transform(x0_encoded).ravel()
                
                # 确定连续因子在编码向量中的索引位置
                cont_indices = []
                idx = 0
                for name in self._factor_names:
                    if self._factor_types.get(name) == "categorical":
                        n_cols = len(self._categorical_levels.get(name, [])) - 1
                        idx += n_cols
                    else:
                        if name in continuous_names:
                            cont_indices.append(idx)
                        idx += 1
                
                # 计算连续因子在 scaled 空间的边界
                lower_encoded = np.array([self._encode_factors(build_factors_from_cont(cont_lower))])
                upper_encoded = np.array([self._encode_factors(build_factors_from_cont(cont_upper))])
                lower_scaled = self._scaler_X.transform(lower_encoded).ravel()
                upper_scaled = self._scaler_X.transform(upper_encoded).ravel()
                
                # 构建 scaled 空间的边界（只对连续因子维度优化，类别因子维度固定）
                bounds_scaled = []
                for i in range(self._k):
                    if i in cont_indices:
                        lo = min(lower_scaled[i], upper_scaled[i])
                        hi = max(lower_scaled[i], upper_scaled[i])
                        bounds_scaled.append((lo, hi))
                    else:
                        # 类别因子 one-hot 列: 固定当前值
                        bounds_scaled.append((x0_scaled[i], x0_scaled[i]))
                
                def neg_ei(x_s):
                    mu, sigma = self._gpr.predict(x_s.reshape(1, -1), return_std=True)
                    mu_real = self._scaler_y.inverse_transform(mu.reshape(-1, 1)).ravel()[0]
                    sigma_real = float(sigma[0]) * self._scaler_y.scale_[0]
                    if sigma_real < 1e-8:
                        return 0.0
                    if maximize:
                        z = (mu_real - best_y) / sigma_real
                        ei = (mu_real - best_y) * norm.cdf(z) + sigma_real * norm.pdf(z)
                    else:
                        z = (best_y - mu_real) / sigma_real
                        ei = (best_y - mu_real) * norm.cdf(z) + sigma_real * norm.pdf(z)
                    return -ei

                try:
                    result = minimize(neg_ei, x0_scaled, method='L-BFGS-B', bounds=bounds_scaled)
                    if -result.fun > global_best_val:
                        global_best_val = -result.fun
                        # 从 scaled 空间还原到实际值
                        x_real = self._scaler_X.inverse_transform(result.x.reshape(1, -1)).ravel()
                        decoded_factors = self._decode_row(x_real.tolist())
                        # 类别因子用本轮的固定值覆盖（避免 decode 误差）
                        decoded_factors.update(cat_dict)
                        global_best_factors = decoded_factors
                except Exception:
                    continue

        if global_best_factors is None:
            return json.dumps({"optimal_factors": {}, "predicted_response": 0.0, "prediction_std": 0.0})

        # 用最优因子做最终预测
        x_opt = np.array([self._encode_factors(global_best_factors)])
        x_opt_scaled = self._scaler_X.transform(x_opt)
        mu_opt, sigma_opt = self._gpr.predict(x_opt_scaled, return_std=True)
        mu_real = self._scaler_y.inverse_transform(mu_opt.reshape(-1, 1)).ravel()[0]
        sigma_real = float(sigma_opt[0]) * self._scaler_y.scale_[0]

        # 连续因子保留4位小数，类别因子保留字符串
        optimal_factors = {}
        for name in self._factor_names:
            val = global_best_factors.get(name, 0.0)
            if self._factor_types.get(name) == "categorical":
                optimal_factors[name] = str(val)
            else:
                optimal_factors[name] = round(float(val), 4)

        return json.dumps({
            "optimal_factors": optimal_factors,
            "predicted_response": round(float(mu_real), 4),
            "prediction_std": round(float(sigma_real), 4),
            "direction": "maximize" if maximize else "minimize"
        }, ensure_ascii=False)

    # ══════════════ 智能停止 ══════════════

    def should_stop(self, remaining_json: str, best_observed: float,
                    maximize: bool = True, ei_threshold: float = 0.01,
                    min_runs_ratio: float = 1.5) -> str:
        """
        判断是否应该提前停止实验。
        
        使用 Expected Improvement (EI) 准则:
          如果剩余所有组的最大 EI < ei_threshold × |best_observed|，建议停止。
        
        remaining_json: '[{"温度": 90, ...}, ...]' — 剩余待执行的因子组合
        best_observed: 当前已观测到的最优响应值
        maximize: True=最大化目标, False=最小化目标
        ei_threshold: EI 相对阈值（占 |best_observed| 的比例），默认 1%
        min_runs_ratio: 最少执行的数据量相对 min_samples 的倍数，默认 1.5 倍
        
        返回 JSON: {"should_stop": true/false, "reason": "...",
                     "best_predicted": 91.2, "best_predicted_std": 2.1,
                     "max_ei": 0.003}
        """
        if not self._is_active or self._gpr is None:
            return json.dumps({"should_stop": False, "reason": "模型未激活",
                               "best_predicted": 0.0, "best_predicted_std": 0.0,
                               "max_ei": 0.0})

        remaining = json.loads(remaining_json)
        if len(remaining) == 0:
            return json.dumps({"should_stop": True, "reason": "没有剩余实验",
                               "best_predicted": 0.0, "best_predicted_std": 0.0,
                               "max_ei": 0.0})

        # ★ 修复 (v3): 使用 _encode_factors() 编码，正确处理类别因子的 one-hot
        # 原来的 bug: 直接用 factors.get(name, 0.0) 构建数组，类别因子的字符串值混入浮点数组
        X_remaining = np.array([self._encode_factors(factors) for factors in remaining])
        X_remaining_scaled = self._scaler_X.transform(X_remaining)
        
        # 计算 EI
        ei_values = self._expected_improvement(X_remaining_scaled, best_observed, maximize)
        max_ei = float(np.max(ei_values))
        
        # 同时获取最佳预测值
        y_pred_scaled, y_std_scaled = self._gpr.predict(X_remaining_scaled, return_std=True)
        y_pred = self._scaler_y.inverse_transform(y_pred_scaled.reshape(-1, 1)).ravel()
        y_std = y_std_scaled * self._scaler_y.scale_[0]
        
        if maximize:
            best_idx = np.argmax(y_pred)
        else:
            best_idx = np.argmin(y_pred)
        best_mean = float(y_pred[best_idx])
        best_std_val = float(y_std[best_idx])
        
        # 停止条件:
        #   1. 已积累足够数据 (>= min_samples × min_runs_ratio)
        #   2. 所有剩余组的最大 EI 低于阈值
        abs_threshold = ei_threshold * max(abs(best_observed), 1e-6)
        has_enough_data = len(self._y_raw) >= int(self._min_samples * min_runs_ratio)
        ei_is_low = max_ei < abs_threshold
        
        should_stop = has_enough_data and ei_is_low

        reason = ""
        if should_stop:
            reason = (f"剩余 {len(remaining)} 组实验的最大 Expected Improvement = {max_ei:.4f}，"
                     f"低于阈值 {abs_threshold:.4f} (= {ei_threshold*100:.1f}% × |{best_observed:.2f}|)，"
                     f"继续实验预计收益极小")

        return json.dumps({
            "should_stop": should_stop,
            "reason": reason,
            "best_predicted": round(best_mean, 4),
            "best_predicted_std": round(best_std_val, 4),
            "max_ei": round(max_ei, 6)
        }, ensure_ascii=False)

    # ══════════════ 敏感性分析 ══════════════

    def get_sensitivity(self) -> str:
        """
        获取参数敏感性（基于 lengthscale 倒数）。
        类别因子的 one-hot 列合并为一个敏感性值（取最大倒数值）
        返回 JSON: {"温度": 0.85, "压力": 0.62, ...}（已归一化到 0-1）
        """
        if not self._is_active or self._gpr is None:
            return json.dumps({})

        try:
            params = self._gpr.kernel_.get_params()
            ls = params.get("k1__k2__length_scale", [])

            if not hasattr(ls, '__iter__'):
                ls = [ls]

            inv_ls = [1.0 / float(l) for l in ls]
            
            # 映射回原始因子名（合并 one-hot 列）
            raw_sensitivity = {}
            idx = 0
            for name in self._factor_names:
                if self._factor_types.get(name) == "categorical":
                    n_cols = len(self._categorical_levels.get(name, [])) - 1
                    if n_cols > 0 and idx + n_cols <= len(inv_ls):
                        # 类别因子: 取其所有 one-hot 列的最大 inv_ls
                        raw_sensitivity[name] = max(inv_ls[idx:idx + n_cols])
                        idx += n_cols
                    else:
                        idx += max(n_cols, 0)
                else:
                    if idx < len(inv_ls):
                        raw_sensitivity[name] = inv_ls[idx]
                        idx += 1
            
            max_inv = max(raw_sensitivity.values()) if raw_sensitivity else 1.0
            sensitivity = {k: round(v / max_inv, 4) for k, v in raw_sensitivity.items()}
            sorted_sens = dict(sorted(sensitivity.items(), key=lambda x: -x[1]))
            return json.dumps(sorted_sens, ensure_ascii=False)
        except Exception:
            return json.dumps({})

    # ══════════════ 序列化 ══════════════

    _SERIALIZE_VERSION = 3  # ★ 修复 (v3): 版本升级

    def serialize(self) -> bytes:
        """
        将整个模型状态序列化为 bytes（供 C# 存入 MySQL LONGBLOB）。
        
        训练数据与 sklearn 模型对象分开存储。
        数据部分 (JSON-safe) 即使 sklearn 版本升级也不受影响。
        """
        state = {
            "_version": self._SERIALIZE_VERSION,
            # ── JSON-safe 部分: 永远可恢复 ──
            "factor_names": self._factor_names,
            "bounds": self._bounds,
            "min_samples": self._min_samples,
            "X_raw": self._X_raw,
            "y_raw": self._y_raw,
            "sources": self._sources,
            "batch_names": self._batch_names,
            "timestamps": self._timestamps,
            "evolution": self._evolution,
            "is_active": self._is_active,
            # 类别因子元数据
            "factor_types": self._factor_types,
            "categorical_levels": self._categorical_levels,
            "k_original": self._k_original,
            "gpr_feature_names": self._gpr_feature_names,
            # ── Pickle 部分: 依赖 sklearn 版本 ──
            "scaler_X": self._scaler_X,
            "scaler_y": self._scaler_y,
            "gpr": self._gpr,
        }
        return pickle.dumps(state)

    def deserialize(self, data: bytes) -> None:
        """
        从 bytes 恢复模型状态。
        
        如果 sklearn 模型对象反序列化失败（版本不兼容），
        自动用训练数据重新训练模型，而不是直接崩溃。
        """
        state = pickle.loads(data)
        
        # ── 恢复 JSON-safe 部分（永远成功）──
        self._factor_names = state["factor_names"]
        self._bounds = state["bounds"]
        self._min_samples = state["min_samples"]
        self._X_raw = state["X_raw"]
        self._y_raw = state["y_raw"]
        self._sources = state.get("sources", ["unknown"] * len(self._y_raw))
        self._batch_names = state.get("batch_names", [""] * len(self._y_raw))
        self._timestamps = state.get("timestamps", [""] * len(self._y_raw))
        self._evolution = state.get("evolution", [])
        
        # 恢复类别因子元数据（兼容旧数据：全部视为连续因子）
        self._factor_types = state.get("factor_types", 
            {name: "continuous" for name in self._factor_names})
        self._categorical_levels = state.get("categorical_levels", {})
        self._k_original = state.get("k_original", len(self._factor_names))
        self._gpr_feature_names = state.get("gpr_feature_names", list(self._factor_names))
        self._k = len(self._gpr_feature_names)
        
        # ── 尝试恢复 sklearn 对象 ──
        try:
            self._scaler_X = state["scaler_X"]
            self._scaler_y = state["scaler_y"]
            self._gpr = state["gpr"]
            self._is_active = state["is_active"]
            
            # 验证模型确实能用（防止 partial corruption）
            if self._is_active and self._gpr is not None:
                test_x = np.zeros((1, self._k))
                self._scaler_X.transform(test_x)  # 简单验证
                
        except Exception:
            # sklearn 版本不兼容或 pickle 数据损坏
            # → 用训练数据自动重新训练
            self._gpr = None
            self._is_active = False
            self._scaler_X = StandardScaler()
            self._scaler_y = StandardScaler()
            
            if len(self._y_raw) >= self._min_samples:
                self.train()  # 自动重新训练

    def get_evolution_data(self) -> str:
        """获取模型演进历史（R²/RMSE 随数据量的变化）"""
        return json.dumps(self._evolution, ensure_ascii=False)

    def get_training_data(self) -> str:
        """
        获取当前训练数据集（供持久化到数据库 + UI 展示）。
        从 one-hot 编码还原为原始因子值（包括类别因子标签）
        """
        data = []
        for i in range(len(self._y_raw)):
            factors = self._decode_row(self._X_raw[i])
            data.append({
                "factors": factors,
                "response": self._y_raw[i],
                "source": self._sources[i] if i < len(self._sources) else "unknown",
                "batch_name": self._batch_names[i] if i < len(self._batch_names) else "",
                "timestamp": self._timestamps[i] if i < len(self._timestamps) else ""
            })
        return json.dumps(data, ensure_ascii=False)

    def _decode_row(self, encoded_row: list) -> dict:
        """从 one-hot 编码行还原为原始因子字典"""
        factors = {}
        idx = 0
        for name in self._factor_names:
            if self._factor_types.get(name) == "categorical":
                levels = self._categorical_levels.get(name, [])
                n_cols = max(len(levels) - 1, 0)
                if n_cols > 0 and idx + n_cols <= len(encoded_row):
                    # 找到值为 1.0 的位置，对应的水平标签
                    hot_vals = encoded_row[idx:idx + n_cols]
                    active_idx = -1
                    for j, v in enumerate(hot_vals):
                        if abs(v - 1.0) < 1e-6:
                            active_idx = j
                            break
                    if active_idx >= 0:
                        factors[name] = levels[active_idx + 1]  # +1 因为 drop first
                    else:
                        factors[name] = levels[0]  # 参考水平（全为 0）
                    idx += n_cols
                else:
                    factors[name] = levels[0] if levels else ""
            else:
                if idx < len(encoded_row):
                    factors[name] = encoded_row[idx]
                    idx += 1
                else:
                    factors[name] = 0.0
        return factors

    def get_bounds(self) -> str:
        """获取因子范围 JSON（供 C# 传给 SurfacePlotter）"""
        return json.dumps(self._bounds, ensure_ascii=False)

    def get_factor_names(self) -> str:
        """获取因子名列表 JSON"""
        return json.dumps(self._factor_names, ensure_ascii=False)

    def reset(self, keep_data: bool = False) -> None:
        """重置模型到冷启动状态"""
        if not keep_data:
            self._X_raw.clear()
            self._y_raw.clear()
            self._sources.clear()
            self._batch_names.clear()
            self._timestamps.clear()
        self._gpr = None
        self._is_active = False
        self._scaler_X = StandardScaler()
        self._scaler_y = StandardScaler()
        self._evolution.clear()

    # ══════════════ 内部方法 ══════════════

    def _expected_improvement(self, X_scaled: np.ndarray, best_y: float,
                              maximize: bool = True) -> np.ndarray:
        """
        计算 Expected Improvement。
        
        maximize=True:  EI = (μ - best_y) * Φ(z) + σ * φ(z), z = (μ - best_y) / σ
        maximize=False: EI = (best_y - μ) * Φ(z) + σ * φ(z), z = (best_y - μ) / σ
        """
        mu, sigma = self._gpr.predict(X_scaled, return_std=True)
        mu_real = self._scaler_y.inverse_transform(mu.reshape(-1, 1)).ravel()
        sigma_real = sigma * self._scaler_y.scale_[0]

        with np.errstate(divide='ignore', invalid='ignore'):
            if maximize:
                improvement = mu_real - best_y
            else:
                improvement = best_y - mu_real
            z = improvement / sigma_real
            ei = improvement * norm.cdf(z) + sigma_real * norm.pdf(z)
            ei[sigma_real < 1e-8] = 0.0

        return ei


# ═══════════════════════════════════════════════════════
# 模块入口 — 供 C# pythonnet 调用
# ═══════════════════════════════════════════════════════

def create_model(factor_names_json: str, bounds_json: str, min_samples: int = 6,
                 factor_types_json: str = "{}") -> DOEGPRModel:
    """创建 GPR 模型实例"""
    return DOEGPRModel(factor_names_json, bounds_json, min_samples, factor_types_json)

def load_model(serialized_bytes: bytes) -> DOEGPRModel:
    """从序列化数据恢复模型"""
    model = DOEGPRModel('[]', '{}')
    model.deserialize(serialized_bytes)
    return model


# ═══════════════════════════════════════════════════════
# 本地测试
# ═══════════════════════════════════════════════════════

if __name__ == "__main__":
    import random

    print("=== DOE GPR 模型测试 ===\n")

    # ── 测试1: 纯连续因子 ──
    print("--- 测试1: 纯连续因子 ---")
    factors_json = json.dumps(["温度", "压力", "催化剂"])
    bounds_json = json.dumps({"温度": [80, 160], "压力": [10, 30], "催化剂": [1.0, 5.0]})

    model = DOEGPRModel(factors_json, bounds_json, min_samples=6)

    test_data = [
        ({"温度": 80, "压力": 10, "催化剂": 1.0}, 45.2),
        ({"温度": 120, "压力": 20, "催化剂": 3.0}, 78.5),
        ({"温度": 160, "压力": 30, "催化剂": 5.0}, 92.1),
        ({"温度": 100, "压力": 15, "催化剂": 2.0}, 62.3),
        ({"温度": 140, "压力": 25, "催化剂": 4.0}, 85.7),
        ({"温度": 80, "压力": 30, "催化剂": 3.0}, 58.9),
        ({"温度": 160, "压力": 10, "催化剂": 2.0}, 71.4),
        ({"温度": 120, "压力": 20, "催化剂": 5.0}, 88.3),
    ]

    for i, (factors, response) in enumerate(test_data):
        model.append_data(json.dumps(factors), response, "measured",
                         f"TEST_BATCH_{i//4}", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    train_result = json.loads(model.train())
    print(f"  训练: active={train_result['is_active']}, R²={train_result['r_squared']:.4f}")

    pred = json.loads(model.predict(json.dumps({"温度": 130, "压力": 22, "催化剂": 3.0})))
    print(f"  预测(130,22,3.0): mean={pred['mean']:.2f} ± {pred['std']:.2f}")

    optimal = json.loads(model.find_optimal("EI"))
    print(f"  最优因子: {optimal.get('optimal_factors', {})}")
    print(f"  预测响应: {optimal.get('predicted_response', 0):.2f}")

    # 测试 should_stop
    remaining = [{"温度": 90, "压力": 12, "催化剂": 1.5}]
    stop = json.loads(model.should_stop(json.dumps(remaining), 92.1))
    print(f"  停止判断: should_stop={stop['should_stop']}, max_ei={stop['max_ei']:.6f}")

    # ── 测试2: 混合因子（连续+类别）──
    print("\n--- 测试2: 混合因子（连续+类别）---")
    factors_json2 = json.dumps(["温度", "压力", "催化剂类型"])
    bounds_json2 = json.dumps({"温度": [80, 160], "压力": [10, 30], "催化剂类型": ["A", "B", "C"]})
    types_json2 = json.dumps({"温度": "continuous", "压力": "continuous", "催化剂类型": "categorical"})

    model2 = DOEGPRModel(factors_json2, bounds_json2, min_samples=6, factor_types_json=types_json2)
    print(f"  GPR 维度: k={model2._k}, features={model2._gpr_feature_names}")

    test_data2 = [
        ({"温度": 80,  "压力": 10, "催化剂类型": "A"}, 45.2),
        ({"温度": 120, "压力": 20, "催化剂类型": "B"}, 78.5),
        ({"温度": 160, "压力": 30, "催化剂类型": "C"}, 92.1),
        ({"温度": 100, "压力": 15, "催化剂类型": "A"}, 55.3),
        ({"温度": 140, "压力": 25, "催化剂类型": "B"}, 85.7),
        ({"温度": 80,  "压力": 30, "催化剂类型": "C"}, 68.9),
        ({"温度": 160, "压力": 10, "催化剂类型": "A"}, 61.4),
        ({"温度": 120, "压力": 20, "催化剂类型": "C"}, 88.3),
    ]

    for factors, response in test_data2:
        model2.append_data(json.dumps(factors), response, "measured", "MIXED_BATCH")

    train_result2 = json.loads(model2.train())
    print(f"  训练: active={train_result2['is_active']}, R²={train_result2['r_squared']:.4f}")
    print(f"  lengthscales: {train_result2['lengthscales']}")

    pred2 = json.loads(model2.predict(json.dumps({"温度": 130, "压力": 22, "催化剂类型": "B"})))
    print(f"  预测(130,22,B): mean={pred2['mean']:.2f} ± {pred2['std']:.2f}")

    optimal2 = json.loads(model2.find_optimal("EI"))
    print(f"  最优因子: {optimal2.get('optimal_factors', {})}")
    print(f"  预测响应: {optimal2.get('predicted_response', 0):.2f}")

    # 测试 should_stop (含类别因子)
    remaining2 = [
        {"温度": 90, "压力": 12, "催化剂类型": "A"},
        {"温度": 140, "压力": 25, "催化剂类型": "C"},
    ]
    stop2 = json.loads(model2.should_stop(json.dumps(remaining2), 92.1))
    print(f"  停止判断: should_stop={stop2['should_stop']}, max_ei={stop2['max_ei']:.6f}")

    # 测试敏感性
    sens2 = json.loads(model2.get_sensitivity())
    print(f"  敏感性: {sens2}")

    # 测试序列化/反序列化
    serialized = model2.serialize()
    restored = load_model(serialized)
    pred_restored = json.loads(restored.predict(json.dumps({"温度": 130, "压力": 22, "催化剂类型": "B"})))
    print(f"  恢复后预测: mean={pred_restored['mean']:.2f} ± {pred_restored['std']:.2f}")

    print("\n所有测试通过！")
