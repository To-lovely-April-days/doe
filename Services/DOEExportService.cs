using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using MaxChemical.Logging;
using MaxChemical.Modules.DOE.Data;
using MaxChemical.Modules.DOE.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MaxChemical.Modules.DOE.Services
{
    /// <summary>
    /// DOE Excel 报告导出服务
    /// </summary>
    public class DOEExportService
    {
        private readonly IDOERepository _repository;
        private readonly ILogService _logger;

        static DOEExportService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public DOEExportService(IDOERepository repository, ILogService logger)
        {
            _repository = repository;
            _logger = logger?.ForContext<DOEExportService>() ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// 导出 DOE 批次完整报告
        /// </summary>
        public async Task<string> ExportBatchReportAsync(string batchId, string filePath)
        {
            var batch = await _repository.GetBatchWithDetailsAsync(batchId);
            if (batch == null) throw new InvalidOperationException($"未找到批次: {batchId}");

            using var package = new ExcelPackage();

            CreateSummarySheet(package, batch);
            CreateFactorsSheet(package, batch);
            CreateResultsSheet(package, batch);
            if (batch.StopConditions.Count > 0)
                CreateStopConditionsSheet(package, batch);

            var fileInfo = new FileInfo(filePath);
            await package.SaveAsAsync(fileInfo);

            _logger.LogInformation("DOE 报告导出成功: {Path}", filePath);
            return filePath;
        }

        /// <summary>
        /// 导出多批次对比报告
        /// </summary>
        public async Task<string> ExportComparisonReportAsync(List<string> batchIds, string filePath)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("批次对比");

            int row = 1;
            SetHeader(ws, row, 1, new[] { "批次名称", "设计方法", "总组数", "完成组数", "状态", "创建时间" });
            row++;

            foreach (var batchId in batchIds)
            {
                var batch = await _repository.GetBatchWithDetailsAsync(batchId);
                if (batch == null) continue;

                var completed = batch.Runs.Count(r => r.Status == DOERunStatus.Completed);
                ws.Cells[row, 1].Value = batch.BatchName;
                ws.Cells[row, 2].Value = batch.DesignMethod.ToString();
                ws.Cells[row, 3].Value = batch.Runs.Count;
                ws.Cells[row, 4].Value = completed;
                ws.Cells[row, 5].Value = batch.Status.ToString();
                ws.Cells[row, 6].Value = batch.CreatedTime.ToString("yyyy-MM-dd HH:mm");
                row++;
            }

            ws.Cells.AutoFitColumns();
            var fileInfo = new FileInfo(filePath);
            await package.SaveAsAsync(fileInfo);
            return filePath;
        }

        // ── 摘要页 ──

        private void CreateSummarySheet(ExcelPackage package, DOEBatch batch)
        {
            var ws = package.Workbook.Worksheets.Add("方案摘要");

            int row = 1;
            ws.Cells[row, 1].Value = "DOE 实验报告";
            ws.Cells[row, 1].Style.Font.Size = 16;
            ws.Cells[row, 1].Style.Font.Bold = true;
            row += 2;

            var info = new (string, string)[]
            {
                ("批次ID", batch.BatchId),
                ("批次名称", batch.BatchName),
                ("流程名称", batch.FlowName),
                ("设计方法", batch.DesignMethod.ToString()),
                ("状态", batch.Status.ToString()),
                ("因子数量", batch.Factors.Count.ToString()),
                ("响应变量", string.Join(", ", batch.Responses.Select(r => $"{r.ResponseName}({r.Unit})"))),
                ("总实验组数", batch.Runs.Count.ToString()),
                ("已完成组数", batch.Runs.Count(r => r.Status == DOERunStatus.Completed).ToString()),
                ("创建时间", batch.CreatedTime.ToString("yyyy-MM-dd HH:mm:ss")),
            };

            foreach (var (label, value) in info)
            {
                ws.Cells[row, 1].Value = label;
                ws.Cells[row, 1].Style.Font.Bold = true;
                ws.Cells[row, 2].Value = value;
                row++;
            }

            ws.Column(1).Width = 18;
            ws.Column(2).Width = 40;
        }

        // ── 因子定义页 ──

        private void CreateFactorsSheet(ExcelPackage package, DOEBatch batch)
        {
            var ws = package.Workbook.Worksheets.Add("因子定义");

            int row = 1;
            SetHeader(ws, row, 1, new[] { "因子名称", "来源", "下界", "上界", "水平数", "步长" });
            row++;

            foreach (var f in batch.Factors.OrderBy(f => f.SortOrder))
            {
                ws.Cells[row, 1].Value = f.FactorName;
                ws.Cells[row, 2].Value = f.FactorSource.ToString();
                ws.Cells[row, 3].Value = f.LowerBound;
                ws.Cells[row, 4].Value = f.UpperBound;
                ws.Cells[row, 5].Value = f.LevelCount;
                ws.Cells[row, 6].Value = f.StepSize;
                row++;
            }

            ws.Cells.AutoFitColumns();
        }

        // ── 实验结果页 ──

        private void CreateResultsSheet(ExcelPackage package, DOEBatch batch)
        {
            var ws = package.Workbook.Worksheets.Add("实验结果");

            var factorNames = batch.Factors.OrderBy(f => f.SortOrder).Select(f => f.FactorName).ToList();
            var responseNames = batch.Responses.OrderBy(r => r.SortOrder).Select(r => r.ResponseName).ToList();

            // 表头
            int col = 1;
            var headers = new List<string> { "组号" };
            headers.AddRange(factorNames);
            headers.AddRange(responseNames);
            headers.AddRange(new[] { "来源", "状态", "开始时间", "结束时间" });

            SetHeader(ws, 1, 1, headers.ToArray());

            // 数据
            int row = 2;
            foreach (var run in batch.Runs.OrderBy(r => r.RunIndex))
            {
                col = 1;
                ws.Cells[row, col++].Value = run.RunIndex + 1;

                var factors = JsonConvert.DeserializeObject<Dictionary<string, object>>(run.FactorValuesJson) ?? new();
                foreach (var name in factorNames)
                    ws.Cells[row, col++].Value = factors.TryGetValue(name, out var v) ? v : "";

                var responses = !string.IsNullOrEmpty(run.ResponseValuesJson)
                    ? JsonConvert.DeserializeObject<Dictionary<string, double>>(run.ResponseValuesJson) ?? new()
                    : new Dictionary<string, double>();
                foreach (var name in responseNames)
                    ws.Cells[row, col++].Value = responses.TryGetValue(name, out var v) ? v : (object)"";

                ws.Cells[row, col++].Value = run.DataSource.ToString();
                ws.Cells[row, col++].Value = run.Status.ToString();
                ws.Cells[row, col++].Value = run.StartTime?.ToString("HH:mm:ss") ?? "";
                ws.Cells[row, col++].Value = run.EndTime?.ToString("HH:mm:ss") ?? "";

                // 高亮已完成行
                if (run.Status == DOERunStatus.Completed)
                {
                    using var range = ws.Cells[row, 1, row, col - 1];
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(240, 248, 255));
                }

                row++;
            }

            ws.Cells.AutoFitColumns();
        }

        // ── 停止条件页 ──

        private void CreateStopConditionsSheet(ExcelPackage package, DOEBatch batch)
        {
            var ws = package.Workbook.Worksheets.Add("停止条件");

            SetHeader(ws, 1, 1, new[] { "条件类型", "响应变量", "运算符", "目标值", "逻辑关系" });

            int row = 2;
            foreach (var c in batch.StopConditions)
            {
                ws.Cells[row, 1].Value = c.ConditionType.ToString();
                ws.Cells[row, 2].Value = c.ResponseName ?? "";
                ws.Cells[row, 3].Value = c.Operator;
                ws.Cells[row, 4].Value = c.TargetValue;
                ws.Cells[row, 5].Value = c.LogicGroup.ToString();
                row++;
            }

            ws.Cells.AutoFitColumns();
        }

        // ── 辅助 ──

        private static void SetHeader(ExcelWorksheet ws, int row, int startCol, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cells[row, startCol + i].Value = headers[i];
            }
            using var range = ws.Cells[row, startCol, row, startCol + headers.Length - 1];
            range.Style.Font.Bold = true;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(46, 117, 182));
            range.Style.Font.Color.SetColor(Color.White);
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        // ══════════════  GPR 数据模板生成 ══════════════

        /// <summary>
        /// 生成 GPR 数据导入模板（空 Excel，只有列头）
        /// 列头 = 因子名 + 响应名
        /// </summary>
        public async Task GenerateGPRTemplateAsync(List<string> factorNames, List<string> responseNames, string filePath)
        {
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("GPR数据");

            int col = 1;

            // 因子列头（蓝色背景）
            foreach (var name in factorNames)
            {
                ws.Cells[1, col].Value = name;
                ws.Cells[1, col].Style.Font.Bold = true;
                ws.Cells[1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(46, 117, 182));
                ws.Cells[1, col].Style.Font.Color.SetColor(Color.White);
                ws.Column(col).Width = 18;
                col++;
            }

            // 响应列头（绿色背景，区分因子）
            foreach (var name in responseNames)
            {
                ws.Cells[1, col].Value = name;
                ws.Cells[1, col].Style.Font.Bold = true;
                ws.Cells[1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[1, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(56, 142, 60));
                ws.Cells[1, col].Style.Font.Color.SetColor(Color.White);
                ws.Column(col).Width = 16;
                col++;
            }

            // 冻结第一行
            ws.View.FreezePanes(2, 1);

            // 添加说明 sheet
            var helpWs = package.Workbook.Worksheets.Add("说明");
            helpWs.Cells[1, 1].Value = "GPR 模型数据导入模板";
            helpWs.Cells[1, 1].Style.Font.Bold = true;
            helpWs.Cells[1, 1].Style.Font.Size = 14;
            helpWs.Cells[3, 1].Value = "使用方法:";
            helpWs.Cells[4, 1].Value = "1. 在「GPR数据」工作表中填写实验数据";
            helpWs.Cells[5, 1].Value = "2. 蓝色列头 = 因子（实验输入参数）";
            helpWs.Cells[6, 1].Value = "3. 绿色列头 = 响应（实验结果）";
            helpWs.Cells[7, 1].Value = "4. 所有单元格必须填写数值，不能为空";
            helpWs.Cells[8, 1].Value = "5. 保存后在「模型分析」页面点击「导入数据」";

            helpWs.Cells[10, 1].Value = "因子列表:";
            for (int i = 0; i < factorNames.Count; i++)
                helpWs.Cells[11 + i, 1].Value = $"  {factorNames[i]}";

            helpWs.Cells[12 + factorNames.Count, 1].Value = "响应列表:";
            for (int i = 0; i < responseNames.Count; i++)
                helpWs.Cells[13 + factorNames.Count + i, 1].Value = $"  {responseNames[i]}";

            helpWs.Column(1).Width = 50;

            var fileInfo = new FileInfo(filePath);
            await package.SaveAsAsync(fileInfo);

            _logger.LogInformation("GPR 数据模板已生成: {Path}, 因子={Factors}, 响应={Responses}",
                filePath, string.Join(",", factorNames), string.Join(",", responseNames));
        }
    }
}