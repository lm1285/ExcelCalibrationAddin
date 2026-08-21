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
	private void AlignRulesToMappings()
	{
		var unmatchedRules = new List<MeasurementRule>(_currentRules.Where(rule => rule != null));
		var alignedRules = new List<MeasurementRule>(_currentMappings.Count);
		for (int i = 0; i < _currentMappings.Count; i++)
		{
			TemplateRegionMapping val = _currentMappings[i];
			MeasurementRule val2 = FindMatchingRule(unmatchedRules, val);
			if (val2 != null)
			{
				unmatchedRules.Remove(val2);
			}
			else
			{
				val2 = new MeasurementRule { IsEnabled = false };
			}
			val2.FieldName = (string.IsNullOrWhiteSpace(val.ProjectName) ? val2.FieldName : val.ProjectName);
			val2.FieldAlias = (string.IsNullOrWhiteSpace(val.ProjectName) ? val2.FieldAlias : val.ProjectName);
			val2.TargetRange = TaskPaneModelCloner.CloneRange(val.MeasurementValueRange) ?? val2.TargetRange;
			val2.SetpointSource = TaskPaneModelCloner.BuildParameterSource(val2.SetpointSource, "设定值", val.SetpointValueRange);
			val2.StandardValueSource = TaskPaneModelCloner.BuildParameterSource(val2.StandardValueSource, "标准值", val.StandardValueRange);
			val2.AverageSource = TaskPaneModelCloner.BuildParameterSource(val2.AverageSource, "平均值", val.AverageValueRange);
			val2.ErrorSource = TaskPaneModelCloner.BuildParameterSource(val2.ErrorSource, "误差", val.ErrorValueRange);
			val2.MpeSource = TaskPaneModelCloner.BuildParameterSource(val2.MpeSource, "技术要求", val.TechnicalRequirementRange);
			val2.RangeSource = TaskPaneModelCloner.BuildParameterSource(val2.RangeSource, "量程", val.RangeValueRange);
			val2.UncertaintySource = TaskPaneModelCloner.BuildParameterSource(val2.UncertaintySource, "不确定度", val.UncertaintyRange);
			val2.ResultSource = TaskPaneModelCloner.BuildParameterSource(val2.ResultSource, "结论", val.ResultRange);
			if (val2.TargetRange != null)
			{
				val2.GroupSize = ResolveGroupSize(val2);
			}
			alignedRules.Add(val2);
		}
		_currentRules = alignedRules;
	}

	private static MeasurementRule FindMatchingRule(
		IReadOnlyList<MeasurementRule> rules,
		TemplateRegionMapping mapping)
	{
		if (rules == null || mapping == null)
		{
			return null;
		}

		var projectName = NormalizeProjectName(mapping.ProjectName);
		var nameMatch = rules.FirstOrDefault(rule =>
			string.Equals(
				NormalizeProjectName(string.IsNullOrWhiteSpace(rule.FieldAlias) ? rule.FieldName : rule.FieldAlias),
				projectName,
				StringComparison.Ordinal));
		if (nameMatch != null)
		{
			return nameMatch;
		}

		return rules.FirstOrDefault(rule => RangesEqual(rule.TargetRange, mapping.MeasurementValueRange));
	}

	private static string NormalizeProjectName(string value)
	{
		return new string((value ?? string.Empty)
			.Where(character => !char.IsWhiteSpace(character) &&
				character != '\u3001' &&
				character != ':' &&
				character != '\uFF1A')
			.ToArray())
			.ToUpperInvariant();
	}

	private static bool RangesEqual(CellRange left, CellRange right)
	{
		return left != null &&
			right != null &&
			string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase) &&
			left.StartRow == right.StartRow &&
			left.EndRow == right.EndRow &&
			left.StartColumn == right.StartColumn &&
			left.EndColumn == right.EndColumn;
	}

	private void CaptureAutoStandardValues()
	{
		for (int i = 0; i < _currentRules.Count; i++)
		{
			Dictionary<int, double?> autoStandardValues = _autoStandardValues;
			int key = i;
			MeasurementRule obj = _currentRules[i];
			autoStandardValues[key] = ((obj != null) ? obj.FixedStandardValue : ((double?)null));
		}
	}

	private void SetRangeForField(int rowIndex, string columnName, CellRange range)
	{
		if (rowIndex < 0 || rowIndex >= _currentMappings.Count)
		{
			return;
		}
		TemplateRegionMapping val = _currentMappings[rowIndex];
		if (!_isBinding) _hasUnsavedChanges = true;
		CellRange val2 = TaskPaneModelCloner.CloneRange(range);
		switch (columnName)
		{
		case "Section":
			val.SectionRange = val2;
			if (GetRule(rowIndex)?.TemplateDefinition != null)
			{
				GetRule(rowIndex).TemplateDefinition.SectionRange = TaskPaneModelCloner.CloneRange(val2);
			}
			break;
		case "Standard":
			val.StandardValueRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
			_currentRules[rowIndex].StandardValueSource = source;
			}, "标准值", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.StandardValue, val2);
			break;
		case "Setpoint":
			val.SetpointValueRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
				_currentRules[rowIndex].SetpointSource = source;
			}, "设定值", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.SetpointValue, val2);
			break;
		case "Measurement":
			val.MeasurementValueRange = val2;
			if (GetRule(rowIndex) != null)
			{
		_currentRules[rowIndex].TargetRange = TaskPaneModelCloner.CloneRange(val2);
				_currentRules[rowIndex].WritableCells = new List<CellAddress>();
				_currentRules[rowIndex].GroupSize = ((val2 == null) ? 1 : ResolveGroupSize(_currentRules[rowIndex]));
			}
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.MeasurementValue, val2);
			break;
		case "Average":
			val.AverageValueRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
			_currentRules[rowIndex].AverageSource = source;
			}, "平均值", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.AverageValue, val2);
			break;
		case "Error":
			val.ErrorValueRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
			_currentRules[rowIndex].ErrorSource = source;
			}, "误差", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.ErrorValue, val2);
			break;
		case "Requirement":
			val.TechnicalRequirementRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
			_currentRules[rowIndex].MpeSource = source;
			}, "技术要求", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.TechnicalRequirement, val2);
			break;
		case "Uncertainty":
			val.UncertaintyRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
			_currentRules[rowIndex].UncertaintySource = source;
			}, "不确定度", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.Uncertainty, val2);
			break;
		case "Range":
			val.RangeValueRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
			_currentRules[rowIndex].RangeSource = source;
			}, "量程", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.RangeValue, val2);
			break;
		case "Result":
			val.ResultRange = val2;
			SetParameterSource(rowIndex, delegate(ParameterSource source)
			{
			_currentRules[rowIndex].ResultSource = source;
			}, "结论", val2);
			UpdateTemplateDefinitionRegion(rowIndex, TemplateRegionRole.Result, val2);
			break;
		}
	}

	private void UpdateTemplateDefinitionRegion(int rowIndex, TemplateRegionRole role, CellRange range)
	{
		var definition = GetRule(rowIndex)?.TemplateDefinition;
		var region = definition?.Regions?.FirstOrDefault(item => item != null && item.Role == role);
		if (region != null)
		{
			region.Range = TaskPaneModelCloner.CloneRange(range);
		}
	}

	private void SetParameterSource(int rowIndex, Action<ParameterSource> assign, string name, CellRange range)
	{
		if (rowIndex >= 0 && rowIndex < _currentRules.Count)
		{
			assign((range == null) ? ((ParameterSource)null) : new ParameterSource
			{
				Name = name,
				Range = TaskPaneModelCloner.CloneRange(range)
			});
		}
	}

	private void CopyPreviousItemStructure(int rowIndex)
	{
		if (rowIndex <= 0 || rowIndex >= _currentMappings.Count)
		{
			return;
		}

		TemplateRegionMapping previous = _currentMappings[rowIndex - 1];
		TemplateRegionMapping current = _currentMappings[rowIndex];
		int rowOffset = ResolveStructureRowOffset(previous, current);
		SetRangeForField(rowIndex, ColumnSection, OffsetRange(previous.SectionRange, rowOffset));
		SetRangeForField(rowIndex, ColumnSetpoint, OffsetRange(previous.SetpointValueRange, rowOffset));
		SetRangeForField(rowIndex, ColumnStandard, OffsetRange(previous.StandardValueRange, rowOffset));
		SetRangeForField(rowIndex, ColumnMeasurement, OffsetRange(previous.MeasurementValueRange, rowOffset));
		SetRangeForField(rowIndex, ColumnAverage, OffsetRange(previous.AverageValueRange, rowOffset));
		SetRangeForField(rowIndex, ColumnError, OffsetRange(previous.ErrorValueRange, rowOffset));
		SetRangeForField(rowIndex, ColumnRequirement, OffsetRange(previous.TechnicalRequirementRange, rowOffset));
		SetRangeForField(rowIndex, ColumnUncertainty, OffsetRange(previous.UncertaintyRange, rowOffset));
		SetRangeForField(rowIndex, ColumnRange, OffsetRange(previous.RangeValueRange, rowOffset));
		SetRangeForField(rowIndex, ColumnResult, OffsetRange(previous.ResultRange, rowOffset));
		BindMappings();
		UpdateTemplateLibraryButtons();
		NotifyGenerationStateChanged();
		MessageBox.Show("已复制上一项区域结构。", "复制上一项结构", MessageBoxButtons.OK, MessageBoxIcon.Information);
	}

	private static int ResolveStructureRowOffset(TemplateRegionMapping previous, TemplateRegionMapping current)
	{
		if (previous?.SectionRange != null && current?.SectionRange != null)
		{
			return current.SectionRange.StartRow - previous.SectionRange.StartRow;
		}

		if (previous?.StandardValueRange != null && current?.StandardValueRange != null)
		{
			return current.StandardValueRange.StartRow - previous.StandardValueRange.StartRow;
		}

		return 0;
	}

	private static CellRange OffsetRange(CellRange range, int rowOffset)
	{
		if (range == null)
		{
			return null;
		}

		var clone = TaskPaneModelCloner.CloneRange(range);
		clone.StartRow = Math.Max(1, clone.StartRow + rowOffset);
		clone.EndRow = Math.Max(clone.StartRow, clone.EndRow + rowOffset);
		return clone;
	}

	private bool HasSelectedField()
	{
		return _highlightedRowIndex >= 0 && _highlightedRowIndex < _currentMappings.Count && _highlightedColumnIndex >= 0;
	}

	private void SelectNextField(int rowIndex, string columnName)
	{
		var fields = new[]
		{
			ColumnSetpoint,
			ColumnStandard,
			ColumnMeasurement,
			ColumnError,
			ColumnRequirement,
			ColumnAverage,
			ColumnUncertainty,
			ColumnRange,
			ColumnResult
		};
		int index = Array.IndexOf(fields, columnName);
		if (index < 0)
		{
			return;
		}

		int nextRow = rowIndex;
		int nextIndex = index + 1;
		if (nextIndex >= fields.Length)
		{
			nextIndex = 0;
			nextRow++;
		}

		if (nextRow >= _currentMappings.Count)
		{
			ClearFieldStatusSelection();
			return;
		}

		string nextColumn = fields[nextIndex];
		_highlightedRowIndex = nextRow;
		_highlightedColumnIndex = ColumnIndexFromName(nextColumn);
		Control nextCard = FindFieldStatusCard(nextRow, nextColumn);
		MarkFieldStatusSelected(nextCard);
		UpdateFieldActionBar();
	}

	private Control FindFieldStatusCard(int rowIndex, string columnName)
	{
		foreach (Control card in _fieldStatusCards)
		{
			FieldStatusTag tag = card.Tag as FieldStatusTag;
			if (tag != null && tag.RowIndex == rowIndex && tag.ColumnName == columnName)
			{
				return card;
			}
		}

		return null;
	}

	private bool CanSaveCurrentTemplate()
	{
		TemplateFingerprint currentFingerprint = _currentFingerprint;
		return !_featuresBlocked && !string.IsNullOrWhiteSpace((currentFingerprint != null) ? currentFingerprint.ExactFingerprint : null);
	}

	private void UpdateTemplateLibraryButtons()
	{
		if (_saveTemplateButton != null)
		{
			_saveTemplateButton.Enabled = CanSaveCurrentTemplate();
			_saveTemplateButton.Text = _isEditingSavedTemplate ? "更新模板" : "保存模板";
		}
		if (_saveAsTemplateButton != null)
		{
			_saveAsTemplateButton.Enabled = CanSaveCurrentTemplate();
		}
		if (_deleteSelectedCalibrationItemsButton != null)
		{
			_deleteSelectedCalibrationItemsButton.Enabled = !_featuresBlocked && _selectedCalibrationRows.Count > 0;
		}
		UpdateFieldActionBar();
	}

	private void UpdateFieldActionBar()
	{
		bool hasSelectedField = !_featuresBlocked && HasSelectedField();
		_useSelectionForFieldButton.Enabled = hasSelectedField;
		_clearSelectedFieldRangeButton.Enabled = hasSelectedField;
		_fieldSelectionLabel.Text = hasSelectedField
			? BuildSelectedFieldDescription()
			: "请先选择字段";
	}

	private string BuildSelectedFieldDescription()
	{
		string projectName = _currentMappings[_highlightedRowIndex].ProjectName;
		string itemName = string.IsNullOrWhiteSpace(projectName)
			? "第 " + (_highlightedRowIndex + 1) + " 项"
			: projectName;
		return "当前字段：" + itemName + " / " + FieldLabelFromColumnName(ColumnNameFromIndex(_highlightedColumnIndex));
	}

    }
}
