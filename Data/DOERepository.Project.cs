using MaxChemical.Modules.DOE.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Threading.Tasks;

namespace MaxChemical.Modules.DOE.Data
{
    /// <summary>
    /// ★ 新增: DOERepository 的项目层扩展（partial class）
    /// 
    /// 不修改 DOERepository.cs 中的任何已有代码。
    /// 所有项目相关的数据库操作放在这个 partial class 中。
    /// </summary>
    public partial class DOERepository
    {
        // ══════════════ Project CRUD ══════════════

        public async Task<string> CreateProjectAsync(DOEProject project)
        {
            if (string.IsNullOrEmpty(project.ProjectId))
                project.ProjectId = $"PROJ_{DateTime.Now:yyyyMMdd}_{Guid.NewGuid().ToString("N")[..6]}";

            const string sql = @"
                INSERT INTO doe_projects 
                (project_id, project_name, objective_description, flow_id, flow_name,
                 current_phase, status, created_time, updated_time)
                VALUES 
                (@projectId, @name, @desc, @flowId, @flowName,
                 @phase, @status, NOW(), NOW())";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@projectId", project.ProjectId);
            cmd.Parameters.AddWithValue("@name", project.ProjectName);
            cmd.Parameters.AddWithValue("@desc", project.ObjectiveDescription ?? "");
            cmd.Parameters.AddWithValue("@flowId", (object?)project.FlowId ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@flowName", (object?)project.FlowName ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@phase", project.CurrentPhase.ToString());
            cmd.Parameters.AddWithValue("@status", project.Status.ToString());
            await cmd.ExecuteNonQueryAsync();

            return project.ProjectId;
        }

        public async Task<DOEProject?> GetProjectAsync(string projectId)
        {
            const string sql = "SELECT * FROM doe_projects WHERE project_id=@id";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapProject(reader) : null;
        }

        public async Task<DOEProject?> GetProjectWithDetailsAsync(string projectId)
        {
            var project = await GetProjectAsync(projectId);
            if (project == null) return null;

            project.ProjectFactors = await GetProjectFactorsAsync(projectId);
            project.Batches = await GetBatchesByProjectAsync(projectId);
            project.RoundSummaries = await GetRoundSummariesAsync(projectId);

            return project;
        }

        public async Task<List<DOEProjectSummary>> GetActiveProjectsAsync()
        {
            const string sql = @"
                SELECT p.*, 
                       (SELECT COUNT(*) FROM doe_batches b WHERE b.project_id = p.project_id) AS total_batches
                FROM doe_projects p
                WHERE p.status = 'Active'
                ORDER BY p.updated_time DESC";

            return await QueryProjectSummaries(sql);
        }

        public async Task<List<DOEProjectSummary>> GetAllProjectSummariesAsync(int limit = 50)
        {
            var sql = $@"
                SELECT p.*, 
                       (SELECT COUNT(*) FROM doe_batches b WHERE b.project_id = p.project_id) AS total_batches
                FROM doe_projects p
                ORDER BY p.updated_time DESC
                LIMIT {limit}";

            return await QueryProjectSummaries(sql);
        }

        public async Task UpdateProjectAsync(DOEProject project)
        {
            const string sql = @"
                UPDATE doe_projects SET
                    project_name = @name,
                    objective_description = @desc,
                    current_phase = @phase,
                    status = @status,
                    updated_time = NOW()
                WHERE project_id = @id";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", project.ProjectId);
            cmd.Parameters.AddWithValue("@name", project.ProjectName);
            cmd.Parameters.AddWithValue("@desc", project.ObjectiveDescription ?? "");
            cmd.Parameters.AddWithValue("@phase", project.CurrentPhase.ToString());
            cmd.Parameters.AddWithValue("@status", project.Status.ToString());
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateProjectPhaseAsync(string projectId, DOEProjectPhase phase)
        {
            const string sql = @"
                UPDATE doe_projects SET current_phase = @phase, updated_time = NOW()
                WHERE project_id = @id";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            cmd.Parameters.AddWithValue("@phase", phase.ToString());
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateProjectStatusAsync(string projectId, DOEProjectStatus status)
        {
            const string sql = @"
                UPDATE doe_projects SET status = @status, updated_time = NOW()
                WHERE project_id = @id";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            cmd.Parameters.AddWithValue("@status", status.ToString());
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateProjectBestAsync(string projectId, double bestValue,
            string bestFactorsJson, int totalExperiments, int completedRounds)
        {
            const string sql = @"
                UPDATE doe_projects SET
                    best_response_value = @best,
                    best_factors_json = @factors,
                    total_experiments = @total,
                    completed_rounds = @rounds,
                    updated_time = NOW()
                WHERE project_id = @id";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            cmd.Parameters.AddWithValue("@best", bestValue);
            cmd.Parameters.AddWithValue("@factors", bestFactorsJson);
            cmd.Parameters.AddWithValue("@total", totalExperiments);
            cmd.Parameters.AddWithValue("@rounds", completedRounds);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task DeleteProjectWithChildrenAsync(string projectId)
        {
            // 级联删除由 FK 处理 (project_factors, round_summaries)
            // 批次的 project_id 设为 NULL（不删除批次本身）
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var tx = await conn.BeginTransactionAsync();

            try
            {
                // 解除批次关联
                using (var cmd = new MySqlCommand(
                    "UPDATE doe_batches SET project_id=NULL, round_number=NULL, project_phase=NULL WHERE project_id=@id", conn, tx))
                {
                    cmd.Parameters.AddWithValue("@id", projectId);
                    await cmd.ExecuteNonQueryAsync();
                }

                // 删除项目（级联删除 factors 和 round_summaries）
                using (var cmd = new MySqlCommand("DELETE FROM doe_projects WHERE project_id=@id", conn, tx))
                {
                    cmd.Parameters.AddWithValue("@id", projectId);
                    await cmd.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

        // ══════════════ ProjectFactor CRUD ══════════════

        public async Task SaveProjectFactorsAsync(string projectId, List<DOEProjectFactor> factors)
        {
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var tx = await conn.BeginTransactionAsync();

            try
            {
                // 删除旧因子
                using (var del = new MySqlCommand(
                    "DELETE FROM doe_project_factors WHERE project_id=@id", conn, tx))
                {
                    del.Parameters.AddWithValue("@id", projectId);
                    await del.ExecuteNonQueryAsync();
                }

                // 插入新因子
                const string ins = @"
                    INSERT INTO doe_project_factors
                    (project_id, factor_name, factor_type, factor_status,
                     current_lower_bound, current_upper_bound, category_levels,
                     fixed_value, fixed_category_level, source_node_id, source_param_name,
                     status_reason, status_changed_in_batch_id, bounds_history_json,
                     sort_order, created_time, updated_time)
                    VALUES
                    (@pid, @name, @type, @status,
                     @lower, @upper, @catLevels,
                     @fixedVal, @fixedCat, @nodeId, @paramName,
                     @reason, @changedBatch, @history,
                     @sort, NOW(), NOW())";

                for (int i = 0; i < factors.Count; i++)
                {
                    var f = factors[i];
                    f.ProjectId = projectId;
                    f.SortOrder = i;

                    using var cmd = new MySqlCommand(ins, conn, tx);
                    cmd.Parameters.AddWithValue("@pid", projectId);
                    cmd.Parameters.AddWithValue("@name", f.FactorName);
                    cmd.Parameters.AddWithValue("@type", f.FactorType.ToString());
                    cmd.Parameters.AddWithValue("@status", f.FactorStatus.ToString());
                    cmd.Parameters.AddWithValue("@lower", f.CurrentLowerBound);
                    cmd.Parameters.AddWithValue("@upper", f.CurrentUpperBound);
                    cmd.Parameters.AddWithValue("@catLevels", (object?)f.CategoryLevels ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@fixedVal", (object?)f.FixedValue ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@fixedCat", (object?)f.FixedCategoryLevel ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@nodeId", (object?)f.SourceNodeId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@paramName", (object?)f.SourceParamName ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@reason", (object?)f.StatusReason ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@changedBatch", (object?)f.StatusChangedInBatchId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@history", (object?)f.BoundsHistoryJson ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@sort", f.SortOrder);
                    await cmd.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

        public async Task<List<DOEProjectFactor>> GetProjectFactorsAsync(string projectId)
        {
            const string sql = @"
                SELECT * FROM doe_project_factors
                WHERE project_id = @id
                ORDER BY sort_order";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            using var reader = await cmd.ExecuteReaderAsync();

            var list = new List<DOEProjectFactor>();
            while (await reader.ReadAsync()) list.Add(MapProjectFactor(reader));
            return list;
        }

        public async Task<List<DOEProjectFactor>> GetActiveProjectFactorsAsync(string projectId)
        {
            const string sql = @"
                SELECT * FROM doe_project_factors
                WHERE project_id = @id AND factor_status = 'Active'
                ORDER BY sort_order";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            using var reader = await cmd.ExecuteReaderAsync();

            var list = new List<DOEProjectFactor>();
            while (await reader.ReadAsync()) list.Add(MapProjectFactor(reader));
            return list;
        }

        public async Task UpdateProjectFactorStatusAsync(int factorId, ProjectFactorStatus status,
            string? reason = null, string? batchId = null)
        {
            const string sql = @"
                UPDATE doe_project_factors SET
                    factor_status = @status,
                    status_reason = @reason,
                    status_changed_in_batch_id = @batch,
                    updated_time = NOW()
                WHERE id = @id";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", factorId);
            cmd.Parameters.AddWithValue("@status", status.ToString());
            cmd.Parameters.AddWithValue("@reason", (object?)reason ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@batch", (object?)batchId ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task UpdateProjectFactorBoundsAsync(int factorId, double lower, double upper,
            string boundsHistoryJson)
        {
            const string sql = @"
                UPDATE doe_project_factors SET
                    current_lower_bound = @lower,
                    current_upper_bound = @upper,
                    bounds_history_json = @history,
                    updated_time = NOW()
                WHERE id = @id";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", factorId);
            cmd.Parameters.AddWithValue("@lower", lower);
            cmd.Parameters.AddWithValue("@upper", upper);
            cmd.Parameters.AddWithValue("@history", boundsHistoryJson);
            await cmd.ExecuteNonQueryAsync();
        }

        // ══════════════ RoundSummary CRUD ══════════════

        public async Task SaveRoundSummaryAsync(DOERoundSummary summary)
        {
            const string sql = @"
                INSERT INTO doe_round_summaries
                (project_id, batch_id, round_number, phase, design_method,
                 active_factor_count, run_count,
                 r_squared, r_squared_adj, r_squared_pred, lack_of_fit_p, gpr_r_squared,
                 factor_ranking_json, screened_out_factors,
                 best_response_value, best_factors_json, optimal_boundary_status_json, max_ei,
                 recommendation, recommendation_reason, user_decision, user_notes,
                 created_time)
                VALUES
                (@pid, @bid, @round, @phase, @method,
                 @factorCount, @runCount,
                 @r2, @r2adj, @r2pred, @lof, @gprR2,
                 @ranking, @screened,
                 @bestVal, @bestFactors, @boundary, @ei,
                 @rec, @recReason, @decision, @notes,
                 NOW())";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@pid", summary.ProjectId);
            cmd.Parameters.AddWithValue("@bid", summary.BatchId);
            cmd.Parameters.AddWithValue("@round", summary.RoundNumber);
            cmd.Parameters.AddWithValue("@phase", summary.Phase.ToString());
            cmd.Parameters.AddWithValue("@method", summary.DesignMethod.ToString());
            cmd.Parameters.AddWithValue("@factorCount", summary.ActiveFactorCount);
            cmd.Parameters.AddWithValue("@runCount", summary.RunCount);
            cmd.Parameters.AddWithValue("@r2", (object?)summary.RSquared ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@r2adj", (object?)summary.RSquaredAdj ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@r2pred", (object?)summary.RSquaredPred ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@lof", (object?)summary.LackOfFitP ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@gprR2", (object?)summary.GprRSquared ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ranking", (object?)summary.FactorRankingJson ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@screened", (object?)summary.ScreenedOutFactors ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@bestVal", (object?)summary.BestResponseValue ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@bestFactors", (object?)summary.BestFactorsJson ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@boundary", (object?)summary.OptimalBoundaryStatusJson ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ei", (object?)summary.MaxEI ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@rec", summary.Recommendation.ToString());
            cmd.Parameters.AddWithValue("@recReason", (object?)summary.RecommendationReason ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@decision", (object?)summary.UserDecision?.ToString() ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@notes", (object?)summary.UserNotes ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }

        public async Task<List<DOERoundSummary>> GetRoundSummariesAsync(string projectId)
        {
            const string sql = @"
                SELECT * FROM doe_round_summaries
                WHERE project_id = @id
                ORDER BY round_number";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            using var reader = await cmd.ExecuteReaderAsync();

            var list = new List<DOERoundSummary>();
            while (await reader.ReadAsync()) list.Add(MapRoundSummary(reader));
            return list;
        }

        public async Task<DOERoundSummary?> GetRoundSummaryByBatchAsync(string batchId)
        {
            const string sql = "SELECT * FROM doe_round_summaries WHERE batch_id=@id";
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", batchId);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapRoundSummary(reader) : null;
        }

        public async Task UpdateRoundDecisionAsync(int summaryId, NextStepRecommendation decision, string? notes)
        {
            const string sql = @"
                UPDATE doe_round_summaries SET user_decision=@dec, user_notes=@notes
                WHERE id=@id";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", summaryId);
            cmd.Parameters.AddWithValue("@dec", decision.ToString());
            cmd.Parameters.AddWithValue("@notes", (object?)notes ?? DBNull.Value);
            await cmd.ExecuteNonQueryAsync();
        }

        // ══════════════ 按项目查询 ══════════════

        public async Task<List<DOEBatch>> GetBatchesByProjectAsync(string projectId)
        {
            const string sql = @"
                SELECT * FROM doe_batches
                WHERE project_id = @id
                ORDER BY round_number, created_time";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            using var reader = await cmd.ExecuteReaderAsync();

            var list = new List<DOEBatch>();
            while (await reader.ReadAsync()) list.Add(MapBatch(reader));
            return list;
        }

        public async Task<List<DOERunRecord>> GetAllCompletedRunsByProjectAsync(string projectId)
        {
            const string sql = @"
                SELECT r.* FROM doe_runs r
                INNER JOIN doe_batches b ON r.batch_id = b.batch_id
                WHERE b.project_id = @id AND r.status = 'Completed'
                ORDER BY b.round_number, r.run_index";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@id", projectId);
            using var reader = await cmd.ExecuteReaderAsync();

            var list = new List<DOERunRecord>();
            while (await reader.ReadAsync()) list.Add(MapRun(reader));
            return list;
        }

        public async Task<GPRModelState?> GetGPRModelByProjectAsync(string projectId, string factorSignature)
        {
            const string sql = @"
                SELECT * FROM doe_gpr_models 
                WHERE project_id = @pid AND factor_signature = @sig
                LIMIT 1";

            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@pid", projectId);
            cmd.Parameters.AddWithValue("@sig", factorSignature);
            using var reader = await cmd.ExecuteReaderAsync();
            return await reader.ReadAsync() ? MapGPRModelState(reader) : null;
        }

        // ══════════════ Mapping 辅助 ══════════════

        private static DOEProject MapProject(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            ProjectId = GetStr(r, "project_id"),
            ProjectName = GetStr(r, "project_name"),
            ObjectiveDescription = GetStrN(r, "objective_description") ?? "",
            FlowId = GetStrN(r, "flow_id"),
            FlowName = GetStrN(r, "flow_name"),
            CurrentPhase = Enum.TryParse<DOEProjectPhase>(GetStr(r, "current_phase"), out var ph) ? ph : DOEProjectPhase.Screening,
            Status = Enum.TryParse<DOEProjectStatus>(GetStr(r, "status"), out var st) ? st : DOEProjectStatus.Active,
            BestResponseValue = IsNull(r, "best_response_value") ? null : GetDbl(r, "best_response_value"),
            BestFactorsJson = GetStrN(r, "best_factors_json"),
            TotalExperiments = GetInt(r, "total_experiments"),
            CompletedRounds = GetInt(r, "completed_rounds"),
            CreatedTime = GetDt(r, "created_time"),
            UpdatedTime = GetDt(r, "updated_time")
        };

        private static DOEProjectFactor MapProjectFactor(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            ProjectId = GetStr(r, "project_id"),
            FactorName = GetStr(r, "factor_name"),
            FactorType = Enum.TryParse<DOEFactorType>(GetStr(r, "factor_type"), out var ft) ? ft : DOEFactorType.Continuous,
            FactorStatus = Enum.TryParse<ProjectFactorStatus>(GetStr(r, "factor_status"), out var fs) ? fs : ProjectFactorStatus.Active,
            CurrentLowerBound = GetDbl(r, "current_lower_bound"),
            CurrentUpperBound = GetDbl(r, "current_upper_bound"),
            CategoryLevels = GetStrN(r, "category_levels"),
            FixedValue = IsNull(r, "fixed_value") ? null : GetDbl(r, "fixed_value"),
            FixedCategoryLevel = GetStrN(r, "fixed_category_level"),
            SourceNodeId = GetStrN(r, "source_node_id"),
            SourceParamName = GetStrN(r, "source_param_name"),
            StatusReason = GetStrN(r, "status_reason"),
            StatusChangedInBatchId = GetStrN(r, "status_changed_in_batch_id"),
            BoundsHistoryJson = GetStrN(r, "bounds_history_json"),
            SortOrder = GetInt(r, "sort_order"),
            CreatedTime = GetDt(r, "created_time"),
            UpdatedTime = GetDt(r, "updated_time")
        };

        private static DOERoundSummary MapRoundSummary(DbDataReader r) => new()
        {
            Id = GetInt(r, "id"),
            ProjectId = GetStr(r, "project_id"),
            BatchId = GetStr(r, "batch_id"),
            RoundNumber = GetInt(r, "round_number"),
            Phase = Enum.TryParse<DOEProjectPhase>(GetStr(r, "phase"), out var ph) ? ph : DOEProjectPhase.Screening,
            DesignMethod = Enum.TryParse<DOEDesignMethod>(GetStr(r, "design_method"), out var dm) ? dm : DOEDesignMethod.FullFactorial,
            ActiveFactorCount = GetInt(r, "active_factor_count"),
            RunCount = GetInt(r, "run_count"),
            RSquared = IsNull(r, "r_squared") ? null : GetDbl(r, "r_squared"),
            RSquaredAdj = IsNull(r, "r_squared_adj") ? null : GetDbl(r, "r_squared_adj"),
            RSquaredPred = IsNull(r, "r_squared_pred") ? null : GetDbl(r, "r_squared_pred"),
            LackOfFitP = IsNull(r, "lack_of_fit_p") ? null : GetDbl(r, "lack_of_fit_p"),
            GprRSquared = IsNull(r, "gpr_r_squared") ? null : GetDbl(r, "gpr_r_squared"),
            FactorRankingJson = GetStrN(r, "factor_ranking_json"),
            ScreenedOutFactors = GetStrN(r, "screened_out_factors"),
            BestResponseValue = IsNull(r, "best_response_value") ? null : GetDbl(r, "best_response_value"),
            BestFactorsJson = GetStrN(r, "best_factors_json"),
            OptimalBoundaryStatusJson = GetStrN(r, "optimal_boundary_status_json"),
            MaxEI = IsNull(r, "max_ei") ? null : GetDbl(r, "max_ei"),
            Recommendation = Enum.TryParse<NextStepRecommendation>(GetStrN(r, "recommendation") ?? "", out var rec) ? rec : NextStepRecommendation.UserDecision,
            RecommendationReason = GetStrN(r, "recommendation_reason"),
            UserDecision = Enum.TryParse<NextStepRecommendation>(GetStrN(r, "user_decision") ?? "", out var ud) ? ud : null,
            UserNotes = GetStrN(r, "user_notes"),
            CreatedTime = GetDt(r, "created_time")
        };

        private async Task<List<DOEProjectSummary>> QueryProjectSummaries(string sql)
        {
            using var conn = CreateConnection();
            await conn.OpenAsync();
            using var cmd = new MySqlCommand(sql, conn);
            using var reader = await cmd.ExecuteReaderAsync();

            var list = new List<DOEProjectSummary>();
            while (await reader.ReadAsync())
            {
                list.Add(new DOEProjectSummary
                {
                    ProjectId = GetStr(reader, "project_id"),
                    ProjectName = GetStr(reader, "project_name"),
                    FlowName = GetStrN(reader, "flow_name"),
                    CurrentPhase = Enum.TryParse<DOEProjectPhase>(GetStr(reader, "current_phase"), out var ph) ? ph : DOEProjectPhase.Screening,
                    Status = Enum.TryParse<DOEProjectStatus>(GetStr(reader, "status"), out var st) ? st : DOEProjectStatus.Active,
                    TotalBatches = HasColumn(reader, "total_batches") ? Convert.ToInt32(reader["total_batches"]) : 0,
                    TotalExperiments = GetInt(reader, "total_experiments"),
                    BestResponseValue = IsNull(reader, "best_response_value") ? null : GetDbl(reader, "best_response_value"),
                    CreatedTime = GetDt(reader, "created_time"),
                    UpdatedTime = GetDt(reader, "updated_time")
                });
            }
            return list;
        }
    }
}