using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Vsto.TaskPane
{
    public partial class CalibrationTaskPaneControl
    {
	private static CellRange ResolveRangeForHighlight(TemplateRegionMapping mapping, string columnName)
	{
		switch (columnName)
		{
		case "Section":
		return TaskPaneModelCloner.CloneRange(mapping.SectionRange);
		case "Setpoint":
		return TaskPaneModelCloner.CloneRange(mapping.SetpointValueRange);
		case "Standard":
		return TaskPaneModelCloner.CloneRange(mapping.StandardValueRange);
		case "Measurement":
		return TaskPaneModelCloner.CloneRange(mapping.MeasurementValueRange);
		case "Average":
		return TaskPaneModelCloner.CloneRange(mapping.AverageValueRange);
		case "Error":
		return TaskPaneModelCloner.CloneRange(mapping.ErrorValueRange);
		case "Requirement":
		return TaskPaneModelCloner.CloneRange(mapping.TechnicalRequirementRange);
		case "Uncertainty":
		return TaskPaneModelCloner.CloneRange(mapping.UncertaintyRange);
		case "Range":
		return TaskPaneModelCloner.CloneRange(mapping.RangeValueRange);
		case "Result":
		return TaskPaneModelCloner.CloneRange(mapping.ResultRange);
		case "Project":
		return TaskPaneModelCloner.CloneRange(mapping.SectionRange) ?? TaskPaneModelCloner.CloneRange(mapping.MeasurementValueRange) ?? TaskPaneModelCloner.CloneRange(mapping.StandardValueRange) ?? TaskPaneModelCloner.CloneRange(mapping.SetpointValueRange) ?? TaskPaneModelCloner.CloneRange(mapping.AverageValueRange) ?? TaskPaneModelCloner.CloneRange(mapping.ErrorValueRange) ?? TaskPaneModelCloner.CloneRange(mapping.TechnicalRequirementRange) ?? TaskPaneModelCloner.CloneRange(mapping.UncertaintyRange) ?? TaskPaneModelCloner.CloneRange(mapping.RangeValueRange) ?? TaskPaneModelCloner.CloneRange(mapping.ResultRange);
		default:
			return null;
		}
	}

	private static int ColumnIndexFromName(string columnName)
	{
		switch (columnName)
		{
		case "Project":
			return 0;
		case "Section":
			return 1;
		case "Setpoint":
			return 2;
		case "Standard":
			return 3;
		case "Measurement":
			return 4;
		case "Average":
			return 5;
		case "Error":
			return 6;
		case "Requirement":
			return 7;
		case "Uncertainty":
			return 8;
		case "Range":
			return 9;
		case "Result":
			return 10;
		default:
			return -1;
		}
	}

	private static string FieldLabelFromColumnName(string columnName)
	{
		switch (columnName)
		{
		case "Section":
			return "项目区域";
		case "Setpoint":
			return "设定值";
		case "Standard":
			return "标准值";
		case "Measurement":
			return "测量值";
		case "Average":
			return "平均值";
		case "Error":
			return "误差";
		case "Requirement":
			return "技术要求";
		case "Uncertainty":
			return "不确定度";
		case "Range":
			return "量程";
		case "Result":
			return "结论";
		default:
			return columnName ?? string.Empty;
		}
	}

	private static string ColumnNameFromIndex(int columnIndex)
	{
		switch (columnIndex)
		{
		case 0:
			return "Project";
		case 1:
			return "Section";
		case 2:
			return "Setpoint";
		case 3:
			return "Standard";
		case 4:
			return "Measurement";
		case 5:
			return "Average";
		case 6:
			return "Error";
		case 7:
			return "Requirement";
		case 8:
			return "Uncertainty";
		case 9:
			return "Range";
		case 10:
			return "Result";
		default:
			return string.Empty;
		}
	}

	private static ExcelCalibrationAddin.Core.Models.GenerationConfiguration CloneGenerationConfiguration(ExcelCalibrationAddin.Core.Models.GenerationConfiguration configuration)
	{
		if (configuration == null)
		{
			return null;
		}
		return new ExcelCalibrationAddin.Core.Services.GenerationConfigurationStore().Clone(configuration);
	}

	private static string RangeToDisplay(CellRange range)
	{
		return (range == null) ? string.Empty : ToAddress(range);
	}

	private static string ToAddress(CellRange range)
	{
		string text = ToColumnName(range.StartColumn) + range.StartRow;
		string text2 = ToColumnName(range.EndColumn) + range.EndRow;
		return (text == text2) ? text : (text + ":" + text2);
	}

	private static string ToColumnName(int columnNumber)
	{
		int num = ((columnNumber <= 0) ? 1 : columnNumber);
		string text = string.Empty;
		while (num > 0)
		{
			int num2 = (num - 1) % 26;
			text = (char)(65 + num2) + text;
			num = (num - num2) / 26;
		}
		return text;
	}

	private static string FormatNullableNumber(double? value)
	{
		return value.HasValue ? value.Value.ToString("G15", CultureInfo.CurrentCulture) : string.Empty;
	}

	private static string BuildStandardValueSummary(MeasurementRule rule)
	{
		string standardValue = FormatNullableNumber(rule?.FixedStandardValue);
		return string.IsNullOrWhiteSpace(standardValue)
			? "标准值：未识别"
			: "标准值：" + standardValue;
	}

	private static int CountCells(CellRange range)
	{
		if (range == null || range.EndRow < range.StartRow || range.EndColumn < range.StartColumn)
		{
			return 0;
		}
		return (range.EndRow - range.StartRow + 1) * (range.EndColumn - range.StartColumn + 1);
	}

	private static int ResolveGroupSize(MeasurementRule rule)
	{
		if (rule?.WritableCells != null && rule.WritableCells.Count > 0)
		{
			return Math.Max(1, rule.WritableCells.Count);
		}
		return Math.Max(1, CountCells(rule?.TargetRange));
	}

    }
}
