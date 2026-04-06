using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Data.Services;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Models;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;

namespace MaxChemical.Modules.DOE.Data
{
    /// <summary>
    /// DOE 数据仓储实现 — 基于 MySQL，复用现有 IDatabaseConfigService 获取连接字符串
    /// </summary>
    public partial class DOERepository : IDOERepository
    {
        private readonly IDatabaseConfigService _dbConfig;
        private readonly ILogService _logger;

        public DOERepository(IDatabaseConfigService dbConfig, ILogService logger)
        {
            _dbConfig = dbConfig ?? throw new ArgumentNullException(nameof(dbConfig));
            _logger = logger?.ForContext<DOERepository>() ?? throw new ArgumentNullException(nameof(logger));
        }

        private MySqlConnection CreateConnection()
        {
            return new MySqlConnection(_dbConfig.GetConnectionString());
        }

        // ══════════════ DbDataReader 扩展辅助 ══════════════

        private static string GetStr(DbDataReader r, string col) => r.GetString(r.GetOrdinal(col));
        private static int GetInt(DbDataReader r, string col) => r.GetInt32(r.GetOrdinal(col));
        private static double GetDbl(DbDataReader r, string col) => r.GetDouble(r.GetOrdinal(col));
        private static bool GetBool(DbDataReader r, string col) => r.GetBoolean(r.GetOrdinal(col));
        private static DateTime GetDt(DbDataReader r, string col) => r.GetDateTime(r.GetOrdinal(col));
        private static bool IsNull(DbDataReader r, string col) => r.IsDBNull(r.GetOrdinal(col));
        private static string? GetStrN(DbDataReader r, string col) => IsNull(r, col) ? null : GetStr(r, col);
        private static double? GetDblN(DbDataReader r, string col) => IsNull(r, col) ? null : GetDbl(r, col);
        private static DateTime? GetDtN(DbDataReader r, string col) => IsNull(r, col) ? null : GetDt(r, col);

        // ══════════════ Batch ══════════════

        public async Task<string> CreateBatchAsync(DOEBatch batch)
        {
            batch.BatchId = $"DOE_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString("N")[..6]}";

            const string sql = @"
    INSERT INTO doe_batches (batch_id, flow_id, flow_name, batch_name, design_method, status, design_config_json,
        project_id, round_number, project_phase)
    VALUES (@batchId, @flowId, @flowName, @batchName, @designMethod, @status, @configJson,
        @projectId, @roundNumber, @projectPhase)";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batch.BatchId);
            cmd.Parameters.AddWithValue("@flowId", batch.FlowId);
            cmd.Parameters.AddWithValue("@flowName", batch.FlowName);
            cmd.Parameters.AddWithValue("@batchName", batch.BatchName);
            cmd.Parameters.AddWithValue("@designMethod", batch.DesignMethod.ToString());
            cmd.Parameters.AddWithValue("@status", batch.Status.ToString());
            cmd.Parameters.AddWithValue("@configJson", batch.DesignConfigJson ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@projectId", (object?)batch.ProjectId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@roundNumber", (object?)batch.RoundNumber ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@projectPhase", (object?)batch.ProjectPhase?.ToString() ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync();

            _logger.LogInformation("创建 DOE 批次: {BatchId} - {BatchName}", batch.BatchId, batch.BatchName);
            return batch.BatchId;
        }

        public async Task<DOEBatch?> GetBatchAsync(string batchId)
        {
            const string sql = "SELECT * FROM doe_batches WHERE batch_id = @batchId";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapBatch(reader) : null;
        }

        public async Task<DOEBatch?> GetBatchWithDetailsAsync(string batchId)
        {
            var batch = await GetBatchAsync(batchId);
            if (batch == null) return null;

            batch.Factors = await GetFactorsAsync(batchId);
            batch.Responses = await GetResponsesAsync(batchId);
            batch.Runs = await GetRunsAsync(batchId);
            batch.StopConditions = await GetStopConditionsAsync(batchId);
            return batch;
        }

        public async Task<List<DOEBatch>> GetBatchesByFlowAsync(string flowId)
        {
            const string sql = "SELECT * FROM doe_batches WHERE flow_id = @flowId ORDER BY created_time DESC";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOEBatch>();
            while (await reader.ReadAsync()) list.Add(MapBatch(reader));
            return list;
        }

        public async Task<List<DOEBatch>> GetRecentBatchesAsync(int count = 20)
        {
            var sql = $"SELECT * FROM doe_batches ORDER BY created_time DESC LIMIT {count}";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOEBatch>();
            while (await reader.ReadAsync()) list.Add(MapBatch(reader));
            return list;
        }

        public async Task UpdateBatchStatusAsync(string batchId, DOEBatchStatus status)
        {
            const string sql = "UPDATE doe_batches SET status = @status, updated_time = NOW() WHERE batch_id = @batchId";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            cmd.Parameters.AddWithValue("@status", status.ToString());
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateBatchAsync(DOEBatch batch)
        {
            const string sql = @"
                UPDATE doe_batches SET flow_name=@flowName, batch_name=@batchName, 
                design_method=@designMethod, status=@status, design_config_json=@configJson, updated_time=NOW()
                WHERE batch_id=@batchId";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batch.BatchId);
            cmd.Parameters.AddWithValue("@flowName", batch.FlowName);
            cmd.Parameters.AddWithValue("@batchName", batch.BatchName);
            cmd.Parameters.AddWithValue("@designMethod", batch.DesignMethod.ToString());
            cmd.Parameters.AddWithValue("@status", batch.Status.ToString());
            cmd.Parameters.AddWithValue("@configJson", batch.DesignConfigJson ?? (object)DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task DeleteBatchAsync(string batchId)
        {
            const string sql = "DELETE FROM doe_batches WHERE batch_id = @batchId";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            await cmd.ExecuteNonQueryAsync();
            _logger.LogInformation("删除 DOE 批次: {BatchId}", batchId);
        }

        // ══════════════ Factors ══════════════

        public async Task SaveFactorsAsync(string batchId, List<DOEFactor> factors)
        {
            using var conn = CreateConnection();
            await conn.OpenAsync();

            using (var delCmd = new MySqlCommand("DELETE FROM doe_factors WHERE batch_id = @batchId", conn))
            {
                delCmd.Parameters.AddWithValue("@batchId", batchId);
                await delCmd.ExecuteNonQueryAsync();
            }

            const string sql = @"
                INSERT INTO doe_factors (batch_id, factor_name, factor_source, source_node_id, source_param_name, 
                lower_bound, upper_bound, level_count, factor_type, sort_order, category_levels)
                VALUES (@batchId, @name, @source, @nodeId, @paramName, @lower, @upper, @levels, @type, @order, @catLevels)";

            for (int i = 0; i < factors.Count; i++)
            {
                var f = factors[i];
                f.BatchId = batchId;
                f.SortOrder = i;
                using var cmd = new MySqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@batchId", batchId);
                cmd.Parameters.AddWithValue("@name", f.FactorName);
                cmd.Parameters.AddWithValue("@source", f.FactorSource.ToString());
                cmd.Parameters.AddWithValue("@nodeId", f.SourceNodeId ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@paramName", f.SourceParamName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@lower", f.LowerBound);
                cmd.Parameters.AddWithValue("@upper", f.UpperBound);
                cmd.Parameters.AddWithValue("@levels", f.LevelCount);
                cmd.Parameters.AddWithValue("@type", f.FactorType.ToString());
                cmd.Parameters.AddWithValue("@order", f.SortOrder);
                cmd.Parameters.AddWithValue("@catLevels", f.CategoryLevels ?? (object)DBNull.Value);
                await cmd.ExecuteNonQueryAsync();
            }
        }

        public async Task<List<DOEFactor>> GetFactorsAsync(string batchId)
        {
            const string sql = "SELECT * FROM doe_factors WHERE batch_id = @batchId ORDER BY sort_order";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOEFactor>();
            while (await reader.ReadAsync()) list.Add(MapFactor(reader));
            return list;
        }

        // ══════════════ Responses ══════════════

        public async Task SaveResponsesAsync(string batchId, List<DOEResponse> responses)
        {
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using (var delCmd = new MySqlCommand("DELETE FROM doe_responses WHERE batch_id = @batchId", conn))
            {
                delCmd.Parameters.AddWithValue("@batchId", batchId);
                await delCmd.ExecuteNonQueryAsync();
            }

            const string sql = @"
                INSERT INTO doe_responses (batch_id, response_name, unit, collection_method, sort_order)
                VALUES (@batchId, @name, @unit, @method, @order)";

            for (int i = 0; i < responses.Count; i++)
            {
                var r = responses[i];
                r.BatchId = batchId;
                r.SortOrder = i;
                using var cmd = new MySqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@batchId", batchId);
                cmd.Parameters.AddWithValue("@name", r.ResponseName);
                cmd.Parameters.AddWithValue("@unit", r.Unit ?? "");
                cmd.Parameters.AddWithValue("@method", r.CollectionMethod.ToString());
                cmd.Parameters.AddWithValue("@order", r.SortOrder);
                await cmd.ExecuteNonQueryAsync();
            }
        }

        public async Task<List<DOEResponse>> GetResponsesAsync(string batchId)
        {
            const string sql = "SELECT * FROM doe_responses WHERE batch_id = @batchId ORDER BY sort_order";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOEResponse>();
            while (await reader.ReadAsync()) list.Add(MapResponse(reader));
            return list;
        }

        // ══════════════ Runs ══════════════

        public async Task SaveRunAsync(DOERunRecord run)
        {
            const string sql = @"
                INSERT INTO doe_runs (batch_id, run_index, factor_values_json, response_values_json, 
                data_source, experiment_id, status, start_time, end_time)
                VALUES (@batchId, @idx, @factors, @responses, @source, @expId, @status, @start, @end)";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", run.BatchId);
            cmd.Parameters.AddWithValue("@idx", run.RunIndex);
            cmd.Parameters.AddWithValue("@factors", run.FactorValuesJson);
            cmd.Parameters.AddWithValue("@responses", run.ResponseValuesJson ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@source", run.DataSource.ToString());
            cmd.Parameters.AddWithValue("@expId", run.ExperimentId ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@status", run.Status.ToString());
            cmd.Parameters.AddWithValue("@start", run.StartTime ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@end", run.EndTime ?? (object)DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task SaveRunsAsync(List<DOERunRecord> runs)
        {
            foreach (var run in runs) await SaveRunAsync(run);
        }

        public async Task UpdateRunAsync(DOERunRecord run)
        {
            const string sql = @"
                UPDATE doe_runs SET response_values_json=@responses, experiment_id=@expId, 
                status=@status, start_time=@start, end_time=@end
                WHERE batch_id=@batchId AND run_index=@idx";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", run.BatchId);
            cmd.Parameters.AddWithValue("@idx", run.RunIndex);
            cmd.Parameters.AddWithValue("@responses", run.ResponseValuesJson ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@expId", run.ExperimentId ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@status", run.Status.ToString());
            cmd.Parameters.AddWithValue("@start", run.StartTime ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@end", run.EndTime ?? (object)DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task<DOERunRecord?> GetRunAsync(string batchId, int runIndex)
        {
            const string sql = "SELECT * FROM doe_runs WHERE batch_id=@batchId AND run_index=@idx";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            cmd.Parameters.AddWithValue("@idx", runIndex);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapRun(reader) : null;
        }

        public async Task<List<DOERunRecord>> GetRunsAsync(string batchId)
        {
            const string sql = "SELECT * FROM doe_runs WHERE batch_id=@batchId ORDER BY run_index";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOERunRecord>();
            while (await reader.ReadAsync()) list.Add(MapRun(reader));
            return list;
        }

        public async Task<List<DOERunRecord>> GetCompletedRunsAsync(string batchId)
        {
            const string sql = "SELECT * FROM doe_runs WHERE batch_id=@batchId AND status='Completed' ORDER BY run_index";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOERunRecord>();
            while (await reader.ReadAsync()) list.Add(MapRun(reader));
            return list;
        }

        // ══════════════ Stop Conditions ══════════════

        public async Task SaveStopConditionsAsync(string batchId, List<DOEStopCondition> conditions)
        {
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using (var delCmd = new MySqlCommand("DELETE FROM doe_stop_conditions WHERE batch_id=@batchId", conn))
            {
                delCmd.Parameters.AddWithValue("@batchId", batchId);
                await delCmd.ExecuteNonQueryAsync();
            }

            const string sql = @"
                INSERT INTO doe_stop_conditions (batch_id, condition_type, response_name, operator, target_value, logic_group)
                VALUES (@batchId, @type, @respName, @op, @target, @logic)";

            foreach (var c in conditions)
            {
                c.BatchId = batchId;
                using var cmd = new MySqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@batchId", batchId);
                cmd.Parameters.AddWithValue("@type", c.ConditionType.ToString());
                cmd.Parameters.AddWithValue("@respName", c.ResponseName ?? (object)DBNull.Value);
                cmd.Parameters.AddWithValue("@op", c.Operator);
                cmd.Parameters.AddWithValue("@target", c.TargetValue);
                cmd.Parameters.AddWithValue("@logic", c.LogicGroup.ToString());
                await cmd.ExecuteNonQueryAsync();
            }
        }

        public async Task<List<DOEStopCondition>> GetStopConditionsAsync(string batchId)
        {
            const string sql = "SELECT * FROM doe_stop_conditions WHERE batch_id=@batchId";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOEStopCondition>();
            while (await reader.ReadAsync()) list.Add(MapStopCondition(reader));
            return list;
        }

        // ══════════════ GPR Model State ══════════════

        public async Task SaveGPRModelStateAsync(GPRModelState state)
        {
            const string sql = @"
        INSERT INTO doe_gpr_models (flow_id,project_id, factor_signature, model_name, model_state, training_data_json, hyperparams_json,
        data_count, r_squared, rmse, is_active, evolution_history_json, last_trained_time)
        VALUES (@flowId,@projectId, @sig, @name, @model, @data, @hyper, @count, @r2, @rmse, @active, @evo, @trained)
        ON DUPLICATE KEY UPDATE 
        model_name=@name, model_state=@model, training_data_json=@data, hyperparams_json=@hyper,
        data_count=@count, r_squared=@r2, rmse=@rmse, is_active=@active, 
        evolution_history_json=@evo, last_trained_time=@trained, updated_time=NOW()";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", state.FlowId);
            cmd.Parameters.AddWithValue("@projectId", (object?)state.ProjectId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@sig", state.FactorSignature);
            cmd.Parameters.AddWithValue("@name", state.ModelName);
            cmd.Parameters.AddWithValue("@model", state.ModelStateBytes ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@data", state.TrainingDataJson ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@hyper", state.HyperparamsJson ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@count", state.DataCount);
            cmd.Parameters.AddWithValue("@r2", state.RSquared ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@rmse", state.RMSE ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@active", state.IsActive);
            cmd.Parameters.AddWithValue("@evo", state.EvolutionHistoryJson ?? (object)DBNull.Value);
            cmd.Parameters.AddWithValue("@trained", state.LastTrainedTime ?? (object)DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task<GPRModelState?> GetGPRModelStateAsync(string flowId)
        {
            const string sql = "SELECT * FROM doe_gpr_models WHERE flow_id=@flowId";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapGPRModelState(reader) : null;
        }

        public async Task UpdateGPRModelStateAsync(GPRModelState state)
        {
            await SaveGPRModelStateAsync(state);
        }

        public async Task DeleteGPRModelStateAsync(string flowId)
        {
            const string sql = "DELETE FROM doe_gpr_models WHERE flow_id=@flowId";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            await cmd.ExecuteNonQueryAsync();
        }

        // ══════════════ 复合查询 ══════════════

        public async Task<List<DOEBatchSummary>> GetBatchSummariesByFlowAsync(string flowId)
        {
            const string sql = @"
                SELECT b.batch_id, b.batch_name, b.flow_name, b.design_method, b.status, b.created_time,
                       COUNT(r.id) AS total_runs,
                       SUM(CASE WHEN r.status='Completed' THEN 1 ELSE 0 END) AS completed_runs
                FROM doe_batches b
                LEFT JOIN doe_runs r ON b.batch_id = r.batch_id
                WHERE b.flow_id = @flowId
                GROUP BY b.batch_id, b.batch_name, b.flow_name, b.design_method, b.status, b.created_time
                ORDER BY b.created_time DESC";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOEBatchSummary>();
            while (await reader.ReadAsync())
            {
                list.Add(new DOEBatchSummary
                {
                    BatchId = GetStr(reader, "batch_id"),
                    BatchName = GetStr(reader, "batch_name"),
                    FlowName = GetStr(reader, "flow_name"),
                    DesignMethod = Enum.TryParse<DOEDesignMethod>(GetStr(reader, "design_method"), out var m) ? m : DOEDesignMethod.FullFactorial,
                    Status = Enum.TryParse<DOEBatchStatus>(GetStr(reader, "status"), out var s) ? s : DOEBatchStatus.Designing,
                    TotalRuns = Convert.ToInt32(reader[reader.GetOrdinal("total_runs")]),
                    CompletedRuns = Convert.ToInt32(reader[reader.GetOrdinal("completed_runs")]),
                    CreatedTime = GetDt(reader, "created_time")
                });
            }
            return list;
        }

        public async Task<List<DOERunRecord>> GetAllTrainingDataByFlowAsync(string flowId)
        {
            const string sql = @"
                SELECT r.* FROM doe_runs r
                INNER JOIN doe_batches b ON r.batch_id = b.batch_id
                WHERE b.flow_id = @flowId AND r.status = 'Completed'
                ORDER BY r.start_time";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOERunRecord>();
            while (await reader.ReadAsync()) list.Add(MapRun(reader));
            return list;
        }

        /// <summary>
        ///  获取所有批次摘要（不限流程）
        /// </summary>
        public async Task<List<DOEBatchSummary>> GetAllBatchSummariesAsync(int limit = 50)
        {
            var sql = $@"
                SELECT b.batch_id, b.batch_name, b.flow_name, b.design_method, b.status, b.created_time,
                       COUNT(r.id) AS total_runs,
                       SUM(CASE WHEN r.status='Completed' THEN 1 ELSE 0 END) AS completed_runs
                FROM doe_batches b
                LEFT JOIN doe_runs r ON b.batch_id = r.batch_id
                GROUP BY b.batch_id, b.batch_name, b.flow_name, b.design_method, b.status, b.created_time
                ORDER BY b.created_time DESC
                LIMIT {limit}";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<DOEBatchSummary>();
            while (await reader.ReadAsync())
            {
                list.Add(new DOEBatchSummary
                {
                    BatchId = GetStr(reader, "batch_id"),
                    BatchName = GetStr(reader, "batch_name"),
                    FlowName = GetStr(reader, "flow_name"),
                    DesignMethod = Enum.TryParse<DOEDesignMethod>(GetStr(reader, "design_method"), out var m) ? m : DOEDesignMethod.FullFactorial,
                    Status = Enum.TryParse<DOEBatchStatus>(GetStr(reader, "status"), out var s) ? s : DOEBatchStatus.Designing,
                    TotalRuns = Convert.ToInt32(reader[reader.GetOrdinal("total_runs")]),
                    CompletedRuns = Convert.ToInt32(reader[reader.GetOrdinal("completed_runs")]),
                    CreatedTime = GetDt(reader, "created_time")
                });
            }
            return list;
        }

        // ══════════════ Mapping（全部使用 DbDataReader）══════════════

        private static DOEBatch MapBatch(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            BatchId = GetStr(r, "batch_id"),
            FlowId = GetStr(r, "flow_id"),
            FlowName = GetStr(r, "flow_name"),
            BatchName = GetStr(r, "batch_name"),
            DesignMethod = Enum.TryParse<DOEDesignMethod>(GetStr(r, "design_method"), out var m) ? m : DOEDesignMethod.FullFactorial,
            Status = Enum.TryParse<DOEBatchStatus>(GetStr(r, "status"), out var s) ? s : DOEBatchStatus.Designing,
            DesignConfigJson = GetStrN(r, "design_config_json"),
            CreatedTime = GetDt(r, "created_time"),
            UpdatedTime = GetDt(r, "updated_time"),
            // ★ 新增: 安全读取项目关联字段
            ProjectId = HasColumn(r, "project_id") ? GetStrN(r, "project_id") : null,
            RoundNumber = HasColumn(r, "round_number") && !IsNull(r, "round_number")
        ? GetInt(r, "round_number") : null,
            ProjectPhase = HasColumn(r, "project_phase") && !IsNull(r, "project_phase")
        ? (Enum.TryParse<DOEProjectPhase>(GetStr(r, "project_phase"), out var pp) ? pp : null)
        : null
        };

        private static DOEFactor MapFactor(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            BatchId = GetStr(r, "batch_id"),
            FactorName = GetStr(r, "factor_name"),
            FactorSource = Enum.TryParse<Infrastructure.DOE.FactorSourceType>(GetStr(r, "factor_source"), out var fs) ? fs : Infrastructure.DOE.FactorSourceType.ParameterOverride,
            SourceNodeId = GetStrN(r, "source_node_id"),
            SourceParamName = GetStrN(r, "source_param_name"),
            LowerBound = GetDbl(r, "lower_bound"),
            UpperBound = GetDbl(r, "upper_bound"),
            LevelCount = GetInt(r, "level_count"),
            // ★ 兼容旧数据: "Discrete" 映射为 Categorical
            FactorType = ParseFactorType(GetStr(r, "factor_type")),
            SortOrder = GetInt(r, "sort_order"),
            // ★ 新增: 类别因子水平标签
            CategoryLevels = HasColumn(r, "category_levels") ? GetStrN(r, "category_levels") : null
        };

        /// <summary>★ 新增: 解析因子类型，兼容旧的 "Discrete" 值</summary>
        private static DOEFactorType ParseFactorType(string value)
        {
            if (string.Equals(value, "Discrete", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "Categorical", StringComparison.OrdinalIgnoreCase))
                return DOEFactorType.Categorical;
            return DOEFactorType.Continuous;
        }

        /// <summary>★ 新增: 安全检查列是否存在（兼容未升级的数据库）</summary>
        private static bool HasColumn(DbDataReader r, string columnName)
        {
            try
            {
                r.GetOrdinal(columnName);
                return !r.IsDBNull(r.GetOrdinal(columnName));
            }
            catch { return false; }
        }

        private static DOEResponse MapResponse(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            BatchId = GetStr(r, "batch_id"),
            ResponseName = GetStr(r, "response_name"),
            Unit = IsNull(r, "unit") ? "" : GetStr(r, "unit"),
            CollectionMethod = Enum.TryParse<DOECollectionMethod>(GetStr(r, "collection_method"), out var cm) ? cm : DOECollectionMethod.Manual,
            SortOrder = GetInt(r, "sort_order")
        };

        private static DOERunRecord MapRun(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            BatchId = GetStr(r, "batch_id"),
            RunIndex = GetInt(r, "run_index"),
            FactorValuesJson = GetStr(r, "factor_values_json"),
            ResponseValuesJson = GetStrN(r, "response_values_json"),
            DataSource = Enum.TryParse<DOEDataSource>(GetStr(r, "data_source"), out var ds) ? ds : DOEDataSource.Measured,
            ExperimentId = GetStrN(r, "experiment_id"),
            Status = Enum.TryParse<DOERunStatus>(GetStr(r, "status"), out var rs) ? rs : DOERunStatus.Pending,
            StartTime = GetDtN(r, "start_time"),
            EndTime = GetDtN(r, "end_time")
        };

        private static DOEStopCondition MapStopCondition(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            BatchId = GetStr(r, "batch_id"),
            ConditionType = Enum.TryParse<DOEStopConditionType>(GetStr(r, "condition_type"), out var ct) ? ct : DOEStopConditionType.Threshold,
            ResponseName = GetStrN(r, "response_name"),
            Operator = GetStr(r, "operator"),
            TargetValue = GetDbl(r, "target_value"),
            LogicGroup = Enum.TryParse<DOELogicGroup>(GetStr(r, "logic_group"), out var lg) ? lg : DOELogicGroup.And
        };

        private static GPRModelState MapGPRModelState(DbDataReader r)
        {
            var state = new GPRModelState
            {
                Id = GetInt(r, "id"),
                FlowId = GetStr(r, "flow_id"),
                FactorSignature = IsNull(r, "factor_signature") ? "" : GetStr(r, "factor_signature"),
                ModelName = IsNull(r, "model_name") ? "" : GetStr(r, "model_name"),
                DataCount = GetInt(r, "data_count"),
                IsActive = GetBool(r, "is_active"),
                CreatedTime = GetDt(r, "created_time"),
                UpdatedTime = GetDt(r, "updated_time")
            };

            var modelOrd = r.GetOrdinal("model_state");
            if (!r.IsDBNull(modelOrd))
            {
                var length = r.GetBytes(modelOrd, 0, null!, 0, 0);
                state.ModelStateBytes = new byte[length];
                r.GetBytes(modelOrd, 0, state.ModelStateBytes, 0, (int)length);
            }

            state.TrainingDataJson = GetStrN(r, "training_data_json");
            state.HyperparamsJson = GetStrN(r, "hyperparams_json");
            state.RSquared = GetDblN(r, "r_squared");
            state.RMSE = GetDblN(r, "rmse");
            state.EvolutionHistoryJson = GetStrN(r, "evolution_history_json");
            state.LastTrainedTime = GetDtN(r, "last_trained_time");
            state.ProjectId = HasColumn(r, "project_id") ? GetStrN(r, "project_id") : null;
            return state;
        }

        public async Task<GPRModelState?> GetGPRModelStateAsync(string flowId, string factorSignature)
        {
            const string sql = "SELECT * FROM doe_gpr_models WHERE flow_id=@flowId AND factor_signature=@sig";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            cmd.Parameters.AddWithValue("@sig", factorSignature);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapGPRModelState(reader) : null;
        }

        public async Task<List<GPRModelState>> GetGPRModelsByFlowAsync(string flowId)
        {
            const string sql = "SELECT * FROM doe_gpr_models WHERE flow_id=@flowId ORDER BY updated_time DESC";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<GPRModelState>();
            while (await reader.ReadAsync()) list.Add(MapGPRModelState(reader));
            return list;
        }

        public async Task DeleteGPRModelAsync(int modelId)
        {
            const string sql = "DELETE FROM doe_gpr_models WHERE id=@id";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", modelId);
            await cmd.ExecuteNonQueryAsync();
        }

        /// <summary>
        /// 查询同 FlowId + FactorSignature 下有多少个方案
        /// </summary>
        public async Task<int> GetBatchCountBySignatureAsync(string flowId, string factorSignature)
        {
            const string sql = @"
                SELECT COUNT(DISTINCT b.batch_id) 
                FROM doe_batches b
                INNER JOIN doe_factors f ON b.batch_id = f.batch_id
                WHERE b.flow_id = @flowId
                GROUP BY b.batch_id
                HAVING GROUP_CONCAT(f.factor_name ORDER BY f.factor_name SEPARATOR ',') = @sig";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@flowId", flowId);
            cmd.Parameters.AddWithValue("@sig", factorSignature);
            var result = await cmd.ExecuteScalarAsync();
            return result != null ? Convert.ToInt32(result) : 0;
        }

        /// <summary>
        /// 删除批次及所有子表数据
        /// </summary>
        public async Task DeleteBatchWithChildrenAsync(string batchId)
        {
            using var conn = CreateConnection();
            await conn.OpenAsync();

            // 按依赖顺序删除子表
            var tables = new[] { "doe_stop_conditions", "doe_runs", "doe_responses", "doe_factors", "doe_batches" };
            foreach (var table in tables)
            {
                using var cmd = new MySqlCommand($"DELETE FROM {table} WHERE batch_id = @batchId", conn);
                cmd.Parameters.AddWithValue("@batchId", batchId);
                await cmd.ExecuteNonQueryAsync();
            }

            _logger.LogInformation("删除 DOE 批次（含子表）: {BatchId}", batchId);
        }

        public async Task<List<GPRModelState>> GetAllGPRModelsAsync()
        {
            const string sql = "SELECT * FROM doe_gpr_models ORDER BY updated_time DESC";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<GPRModelState>();
            while (await reader.ReadAsync()) list.Add(MapGPRModelState(reader));
            return list;
        }

        public async Task<GPRModelState?> GetGPRModelByIdAsync(int modelId)
        {
            const string sql = "SELECT * FROM doe_gpr_models WHERE id=@id";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", modelId);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapGPRModelState(reader) : null;
        }

        public async Task<List<GPRModelState>> GetGPRModelsByProjectAsync(string projectId)
        {
            const string sql = "SELECT * FROM doe_gpr_models WHERE project_id=@pid ORDER BY updated_time DESC";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@pid", projectId);
            using var reader = await cmd.ExecuteReaderAsync();
            var list = new List<GPRModelState>();
            while (await reader.ReadAsync()) list.Add(MapGPRModelState(reader));
            return list;
        }

    }

    public partial class DOERepository
    {
        // ══════════════  新增: Desirability 配置 ══════════════

        public async Task SaveDesirabilityConfigAsync(string batchId, string configJson)
        {
            const string sql = @"
                INSERT INTO doe_desirability_configs (batch_id, config_json)
                VALUES (@batchId, @configJson)
                ON DUPLICATE KEY UPDATE config_json = @configJson, updated_time = NOW()";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            cmd.Parameters.AddWithValue("@configJson", configJson);
            await cmd.ExecuteNonQueryAsync();

            _logger.LogInformation(" Desirability 配置已保存: {BatchId}", batchId);
        }

        public async Task<string?> GetDesirabilityConfigJsonAsync(string batchId)
        {
            const string sql = "SELECT config_json FROM doe_desirability_configs WHERE batch_id = @batchId";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);

            var result = await cmd.ExecuteScalarAsync();
            return result == null || result == DBNull.Value ? null : result.ToString();
        }

        // ══════════════  新增: OLS 分析结果缓存 ══════════════

        public async Task SaveOlsResultAsync(string batchId, string responseName, string resultJson)
        {
            const string sql = @"
                INSERT INTO doe_ols_results (batch_id, response_name, result_json)
                VALUES (@batchId, @responseName, @resultJson)
                ON DUPLICATE KEY UPDATE result_json = @resultJson, updated_time = NOW()";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            cmd.Parameters.AddWithValue("@responseName", responseName);
            cmd.Parameters.AddWithValue("@resultJson", resultJson);
            await cmd.ExecuteNonQueryAsync();

            _logger.LogInformation(" OLS 分析结果已缓存: {BatchId}/{Response}", batchId, responseName);
        }

        public async Task<string?> GetOlsResultJsonAsync(string batchId, string responseName)
        {
            const string sql = "SELECT result_json FROM doe_ols_results WHERE batch_id = @batchId AND response_name = @responseName";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@batchId", batchId);
            cmd.Parameters.AddWithValue("@responseName", responseName);

            var result = await cmd.ExecuteScalarAsync();
            return result == null || result == DBNull.Value ? null : result.ToString();
        }
    }
}