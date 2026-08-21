using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class CalibrationRibbon
    {
        private const string DefaultRandomRangeTitle = "示值误差随机数生成范围";
        private const string DefaultRandomRangeDetail = "匹配模板后显示系数区间";
        private static readonly Regex NumberRegex = new Regex(@"[-+]?\d+(\.\d+)?([eE][-+]?\d+)?", RegexOptions.Compiled);

        private void CalibrationRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            lblAddinVersion.Label = AddinVersion.DisplayLabel;
            ResetRandomRangeSummary();
            UpdateSampleDataButtonState();
        }

        internal void UpdateSampleDataButtonState()
        {
            var state = Globals.ThisAddIn?.GetCachedGenerationState();
            var enabled = state != null && state.CanGenerate && !state.IsFeatureBlocked && !string.IsNullOrWhiteSpace(state.ExactFingerprint);
            btnSaveSampleData.Enabled = enabled;
            btnViewSampleData.Enabled = enabled;
        }

        private void btnSaveSampleData_Click(object sender, RibbonControlEventArgs e)
        {
            try { Globals.ThisAddIn.ShowSaveSampleData(); } catch (Exception ex) { System.Windows.Forms.MessageBox.Show(ex.Message, "保存样本数据"); }
        }

        private void btnViewSampleData_Click(object sender, RibbonControlEventArgs e)
        {
            try { Globals.ThisAddIn.ShowSampleDataVersions(); } catch (Exception ex) { System.Windows.Forms.MessageBox.Show(ex.Message, "查看样本数据"); }
        }

        internal void UpdateRandomRangeSummary(GenerationConfiguration configuration, System.Collections.Generic.IEnumerable<MeasurementRule> rules)
        {
            var ruleList = (rules ?? Enumerable.Empty<MeasurementRule>())
                .Where(item => item != null)
                .ToList();

            var rule = ruleList
                .FirstOrDefault(item => item != null && item.FixedMpe.HasValue && item.FixedMpe.Value > 0);
            UpdateOverrideRuleItems(ruleList, rule);

            if (rule == null)
            {
                ResetRandomRangeSummary();
                return;
            }

            var itemName = ResolveRuleName(rule);
            if (string.IsNullOrWhiteSpace(itemName))
            {
                itemName = "校准项";
            }

            var rangeDetail = BuildCoefficientRangeText(configuration ?? new GenerationConfiguration(), rule);
            lblRandomRangeTitle.Label = $"{itemName}随机数生成范围";
            lblRandomRangeDetail.Label = rangeDetail;
        }

        internal void ResetRandomRangeSummary()
        {
            lblRandomRangeTitle.Label = DefaultRandomRangeTitle;
            lblRandomRangeDetail.Label = DefaultRandomRangeDetail;
            edtOverrideRange.Text = string.Empty;
            edtOverrideDecimals.Text = string.Empty;
            edtAlarmValue.Text = string.Empty;
            UpdateOverrideRuleItems(Enumerable.Empty<MeasurementRule>(), null);
        }

        internal MeasurementGenerationOverride GetSingleUseOverride()
        {
            var fieldName = (cboOverrideRule.Text ?? string.Empty).Trim();
            var rangeText = (edtOverrideRange.Text ?? string.Empty).Trim();
            var decimalsText = (edtOverrideDecimals.Text ?? string.Empty).Trim();
            var alarmValueText = (edtAlarmValue.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(fieldName) &&
                string.IsNullOrWhiteSpace(rangeText) &&
                string.IsNullOrWhiteSpace(decimalsText) &&
                string.IsNullOrWhiteSpace(alarmValueText))
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(fieldName) &&
                (!string.IsNullOrWhiteSpace(rangeText) || !string.IsNullOrWhiteSpace(decimalsText)))
            {
                throw new InvalidOperationException("请先选择需要临时配置的校准项。");
            }

            var generationOverride = new MeasurementGenerationOverride
            {
                FieldName = fieldName,
                AlarmValue = ParseAlarmValue(alarmValueText)
            };

            if (!string.IsNullOrWhiteSpace(rangeText))
            {
                generationOverride.CoefficientOverride = ParseCoefficientOverride(rangeText);
            }

            if (!string.IsNullOrWhiteSpace(decimalsText))
            {
                int decimalPlaces;
                if (!int.TryParse(decimalsText, NumberStyles.Integer, CultureInfo.CurrentCulture, out decimalPlaces) ||
                    decimalPlaces < 0 ||
                    decimalPlaces > 15)
                {
                    throw new InvalidOperationException("小数位数必须是 0 到 15 之间的整数。");
                }

                generationOverride.DecimalPlaces = decimalPlaces;
            }

            return generationOverride;
        }

        internal void ClearSingleUseOverride()
        {
            cboOverrideRule.Text = string.Empty;
            edtOverrideRange.Text = string.Empty;
            edtOverrideDecimals.Text = string.Empty;
            edtAlarmValue.Text = string.Empty;
        }

        private static double? ParseAlarmValue(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var value = ParseDouble(text.Trim());
            if (!value.HasValue || double.IsNaN(value.Value) || double.IsInfinity(value.Value))
            {
                throw new InvalidOperationException("报警值必须是有效数值。");
            }

            return value.Value;
        }

        private async void btnQuickGenerate_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Trace.WriteLine("[VSTO] Ribbon click: GenerateRandom");
                await Globals.ThisAddIn.GenerateRandomNumbersCurrentWorkbookAsync();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon generate random failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "生成随机数失败");
            }
        }

        private async void btnRecognize_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Trace.WriteLine("[VSTO] Ribbon click: Recognize");
                await Globals.ThisAddIn.RecognizeCurrentWorkbookAsync();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon recognize failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "识别失败");
            }
        }

        private void btnTemplateLibrary_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Trace.WriteLine("[VSTO] Ribbon click: TemplateLibrary");
                Globals.ThisAddIn.ShowTemplateLibraryManager();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon template library failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "模板库管理失败");
            }
        }

        private void btnRandomConfig_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Trace.WriteLine("[VSTO] Ribbon click: RandomConfig");
                Globals.ThisAddIn.ShowRandomGenerationConfiguration();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon random config failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "随机数配置失败");
            }
        }

        private void btnTogglePane_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Trace.WriteLine("[VSTO] Ribbon click: TogglePane");
                Globals.ThisAddIn.ToggleTaskPane();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon toggle pane failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "侧边栏打开失败");
            }
        }

        private void btnSaveMultiArea_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Trace.WriteLine("[VSTO] Ribbon click: SaveMultiAreaTemplate");
                Globals.ThisAddIn.ShowMultiAreaTemplateSave();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon save multi-area template failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "保存多区域模板");
            }
        }

        private void btnRunMultiArea_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Trace.WriteLine("[VSTO] Ribbon click: RunMultiAreaPaste");
                Globals.ThisAddIn.ShowMultiAreaCopyPasteRun();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon run multi-area paste failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "运行多区域粘贴");
            }
        }

        private void btnDeleteMultiArea_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.ShowMultiAreaTemplateDelete();
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"[VSTO] Ribbon delete multi-area template failed: {ex}");
                System.Windows.Forms.MessageBox.Show(ex.Message, "删除多区域模板");
            }
        }

        private static string BuildCoefficientRangeText(GenerationConfiguration configuration, MeasurementRule rule)
        {
            if (!configuration.UseIndependentDeviationControl)
            {
                return FormatCoefficientRange(
                    configuration.UnifiedErrorMinimumCoefficient,
                    configuration.UnifiedErrorMaximumCoefficient);
            }

            return $"负:{FormatCoefficientRange(configuration.NegativeErrorMinimumCoefficient, configuration.NegativeErrorMaximumCoefficient)} " +
                $"正:{FormatCoefficientRange(configuration.PositiveErrorMinimumCoefficient, configuration.PositiveErrorMaximumCoefficient)}";
        }

        private static string FormatCoefficientRange(double minimum, double maximum)
        {
            return $"{FormatCoefficient(minimum)}~{FormatCoefficient(maximum)}";
        }

        private static string FormatCoefficient(double value)
        {
            return value.ToString("0.##", CultureInfo.InvariantCulture);
        }

        private void UpdateOverrideRuleItems(System.Collections.Generic.IEnumerable<MeasurementRule> rules, MeasurementRule preferredRule)
        {
            var currentText = (cboOverrideRule.Text ?? string.Empty).Trim();
            cboOverrideRule.Items.Clear();

            foreach (var name in (rules ?? Enumerable.Empty<MeasurementRule>())
                         .Select(ResolveRuleName)
                         .Where(name => !string.IsNullOrWhiteSpace(name))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = name;
                cboOverrideRule.Items.Add(item);
            }

            if (!string.IsNullOrWhiteSpace(currentText) &&
                cboOverrideRule.Items.Cast<RibbonDropDownItem>().Any(item => string.Equals(item.Label, currentText, StringComparison.OrdinalIgnoreCase)))
            {
                cboOverrideRule.Text = currentText;
            }
            else if (preferredRule != null)
            {
                cboOverrideRule.Text = ResolveRuleName(preferredRule);
            }
            else
            {
                cboOverrideRule.Text = string.Empty;
            }
        }

        private static string ResolveRuleName(MeasurementRule rule)
        {
            if (rule == null)
            {
                return string.Empty;
            }

            var name = string.IsNullOrWhiteSpace(rule.FieldAlias) ? rule.FieldName : rule.FieldAlias;
            return name ?? string.Empty;
        }

        private static MeasurementGenerationCoefficientOverride ParseCoefficientOverride(string text)
        {
            var values = NumberRegex.Matches(text ?? string.Empty)
                .Cast<Match>()
                .Select(match => ParseCoefficient(text, match))
                .Where(value => value.HasValue)
                .Select(value => value.Value)
                .ToList();
            if (values.Count < 2)
            {
                throw new InvalidOperationException("生成区间范围格式无效，请输入 MPE 系数区间，例如 0.2~0.8。");
            }

            var firstRange = NormalizeCoefficientRange(values[0], values[1]);
            var result = new MeasurementGenerationCoefficientOverride
            {
                NegativeMinimumCoefficient = firstRange.minimum,
                NegativeMaximumCoefficient = firstRange.maximum,
                PositiveMinimumCoefficient = firstRange.minimum,
                PositiveMaximumCoefficient = firstRange.maximum,
                AbsoluteMinimumCoefficient = firstRange.minimum,
                AbsoluteMaximumCoefficient = firstRange.maximum
            };

            if (values.Count >= 4)
            {
                var secondRange = NormalizeCoefficientRange(values[2], values[3]);
                result.PositiveMinimumCoefficient = secondRange.minimum;
                result.PositiveMaximumCoefficient = secondRange.maximum;
            }

            return result;
        }

        private static (double minimum, double maximum) NormalizeCoefficientRange(double first, double second)
        {
            var minimum = Math.Min(first, second);
            var maximum = Math.Max(first, second);
            if (minimum < 0 || maximum > 1)
            {
                throw new InvalidOperationException("生成区间系数必须在 0 到 1 之间；如需输入百分比请使用 20%~80%。");
            }

            return (minimum, maximum);
        }

        private static double? ParseCoefficient(string fullText, Match match)
        {
            var value = ParseDouble(match.Value);
            if (!value.HasValue)
            {
                return null;
            }

            var contextStart = Math.Max(0, match.Index);
            var contextLength = Math.Min((fullText ?? string.Empty).Length - contextStart, match.Length + 2);
            var context = (fullText ?? string.Empty).Substring(contextStart, contextLength);
            return context.Contains("%") || context.Contains("％")
                ? value.Value / 100.0
                : value.Value;
        }

        private static double? ParseDouble(string text)
        {
            double value;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
            {
                return value;
            }

            return null;
        }
    }
}
