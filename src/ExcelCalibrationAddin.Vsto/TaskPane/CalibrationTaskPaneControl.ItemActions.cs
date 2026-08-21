using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.Services;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Vsto.TaskPane
{
    public partial class CalibrationTaskPaneControl
    {
	private void RemoveCalibrationItem(int rowIndex)
	{
		if (rowIndex < 0 || rowIndex >= _currentMappings.Count)
		{
			return;
		}
		_currentMappings.RemoveAt(rowIndex);
		if (rowIndex < _currentRules.Count)
		{
			_currentRules.RemoveAt(rowIndex);
		}
		ReindexManualRowsAfterDelete(rowIndex);
		ReindexSelectedRowsAfterDelete(rowIndex);
		ReindexCollapsedRowsAfterDelete(rowIndex);
	}

	private void ReindexCollapsedRowsAfterDelete(int deletedRowIndex)
	{
		List<int> collapsedRows = _collapsedCalibrationRows
			.Where(rowIndex => rowIndex != deletedRowIndex)
			.Select(rowIndex => rowIndex > deletedRowIndex ? rowIndex - 1 : rowIndex)
			.ToList();
		_collapsedCalibrationRows.Clear();
		foreach (int rowIndex in collapsedRows)
		{
			_collapsedCalibrationRows.Add(rowIndex);
		}
	}

	private void AddCalibrationItemButton_Click(object sender, EventArgs e)
	{
		if (_featuresBlocked)
		{
			return;
		}
		int nextIndex = _currentMappings.Count + 1;
		TemplateRegionMapping mapping = new TemplateRegionMapping
		{
			ProjectName = $"新增校准项 {nextIndex}"
		};
		MeasurementRule templateRule = _currentRules.LastOrDefault(item => item != null);
		MeasurementRule rule = templateRule == null ? new MeasurementRule() : TaskPaneModelCloner.CloneRule(templateRule);
		rule.FieldName = mapping.ProjectName;
		rule.FieldAlias = mapping.ProjectName;
		rule.TargetRange = null;
		rule.SetpointSource = null;
		rule.StandardValueSource = null;
		rule.AverageSource = null;
		rule.ErrorSource = null;
		rule.FixedStandardValue = null;
		_currentMappings.Add(mapping);
		_currentRules.Add(rule);
		_manualStandardRows.Add(_currentMappings.Count - 1);
		BindMappings();
		UpdateTemplateLibraryButtons();
		NotifyGenerationStateChanged();
	}

	private async void SaveTemplateButton_Click(object sender, EventArgs e)
	{
		if (_isEditingSavedTemplate)
		{
			SaveEditedTemplate(false);
			return;
		}

		if (!CanSaveCurrentTemplate())
		{
			MessageBox.Show("请先完成模板识别，并确认当前模板可保存。", "保存模板");
			return;
		}
		TemplateSaveDirectory directory = PromptTemplateDirectory();
		if (directory == null)
		{
			return;
		}
		string text = directory.StorageTemplateName;
		bool createNew;
		if (!ConfirmSaveTemplateMode(ref text, out createNew))
		{
			return;
		}
		try
		{
			_saveTemplateButton.Enabled = false;
			UseWaitCursor = true;
			IReadOnlyList<MeasurementRule> savableRules = GetSavableRules();
			TemplateSaveResult saveResult = await Globals.ThisAddIn.SaveCurrentTemplateAsync(
				text,
				TaskPaneModelCloner.CloneFingerprint(_currentFingerprint),
				savableRules,
				CloneGenerationConfiguration(_appliedGenerationConfiguration),
				createNew,
				directory.Metadata);
			_usesTemplateGenerationConfiguration = true;
			_canGenerate = true;
			NotifyGenerationStateChanged();
			_hasUnsavedChanges = false;
			MessageBox.Show(saveResult.Message, "保存模板");
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.Message, "保存模板失败");
		}
		finally
		{
			UseWaitCursor = false;
			UpdateTemplateLibraryButtons();
		}
	}

	private void SaveAsTemplateButton_Click(object sender, EventArgs e)
	{
		SaveEditedTemplate(true);
	}

	private bool SaveEditedTemplate(bool saveAs)
	{
		if (!CanSaveCurrentTemplate()) return false;
		TemplateSaveDirectory directory = saveAs || !_isEditingSavedTemplate
			? PromptTemplateDirectory()
			: new TemplateSaveDirectory { StorageTemplateName = _editingTemplateName, Metadata = _editingDirectoryMetadata };
		if (directory == null) return false;
		try
		{
			_saveTemplateButton.Enabled = false;
			_saveAsTemplateButton.Enabled = false;
			UseWaitCursor = true;
			var saveFingerprint = TaskPaneModelCloner.CloneFingerprint(_currentFingerprint);
			if (_isEditingSavedTemplate && !saveAs && saveFingerprint != null)
			{
				saveFingerprint.ExactFingerprint = _editingTemplateFingerprint;
			}
			var result = Globals.ThisAddIn.SaveCurrentTemplate(
				directory.StorageTemplateName,
				saveFingerprint,
				GetSavableRules(),
				CloneGenerationConfiguration(_appliedGenerationConfiguration),
				!_isEditingSavedTemplate || saveAs,
				directory.Metadata,
				prepareFromWorkbook: !_isEditingSavedTemplate,
				targetRemoteTemplateId: _isEditingSavedTemplate && !saveAs ? _editingRemoteTemplateId : null);
			_hasUnsavedChanges = false;
			if (_isEditingSavedTemplate && saveAs)
			{
				var savedCopy = (Globals.ThisAddIn.GetSavedTemplates() ?? new List<SavedTemplateInfo>())
					.FirstOrDefault(item =>
						string.Equals(item.ExactFingerprint, saveFingerprint.ExactFingerprint, StringComparison.Ordinal) &&
						string.Equals(item.TemplateName, directory.StorageTemplateName, StringComparison.OrdinalIgnoreCase));
				_editingRemoteTemplateId = savedCopy?.RemoteTemplateId ?? string.Empty;
				_editingTemplateName = savedCopy?.TemplateName ?? directory.StorageTemplateName;
				_editingDirectoryMetadata = savedCopy?.DirectoryMetadata ?? directory.Metadata;
				lblRemoteValue.Text = "正在编辑：" + directory.StorageTemplateName;
			}
			MessageBox.Show(result.Message, saveAs ? "另存模板" : "更新模板");
			return true;
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.Message, saveAs ? "另存模板失败" : "更新模板失败");
			return false;
		}
		finally
		{
			UseWaitCursor = false;
			UpdateTemplateLibraryButtons();
		}
	}


	private void UseSelectionForField_Click(object sender, EventArgs e)
	{
		if (!HasSelectedField())
		{
			MessageBox.Show("请先点击一个字段状态卡片。", "修改字段区域");
			return;
		}
		ApplyActiveSelectionToField(_highlightedRowIndex, ColumnNameFromIndex(_highlightedColumnIndex), true);
	}

	private void ClearSelectedFieldRange_Click(object sender, EventArgs e)
	{
		if (!HasSelectedField())
		{
			MessageBox.Show("请先点击一个字段状态卡片。", "清除字段区域");
			return;
		}
		SetRangeForField(_highlightedRowIndex, ColumnNameFromIndex(_highlightedColumnIndex), null);
		RefreshSelectedFieldStatusCard();
	}

	private void ApplyActiveSelectionToField(int rowIndex, string columnName, bool advanceToNextField)
	{
		CellRange activeSelectionRange = Globals.ThisAddIn.GetActiveSelectionRange();
		if (activeSelectionRange == null)
		{
			MessageBox.Show("请先在 Excel 中选择一个有效区域。", "修改字段区域");
			return;
		}

		SetRangeForField(rowIndex, columnName, activeSelectionRange);
		RefreshSelectedFieldStatusCard();
		ShowRangeQualityMessage(rowIndex, columnName, activeSelectionRange);
		if (advanceToNextField)
		{
			SelectNextField(rowIndex, columnName);
		}
	}

	private void ShowRangeQualityMessage(int rowIndex, string columnName, CellRange range)
	{
		string quality = Globals.ThisAddIn.GetActiveSelectionQualitySummary();
		string alignment = BuildRowAlignmentSummary(rowIndex, columnName, range);
		string message = string.Join(Environment.NewLine, new[] { quality, alignment }.Where(item => !string.IsNullOrWhiteSpace(item)));
		if (!string.IsNullOrWhiteSpace(message))
		{
			MessageBox.Show(message, "区域质量提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}

	private string BuildRowAlignmentSummary(int rowIndex, string columnName, CellRange range)
	{
		if (range == null || rowIndex < 0 || rowIndex >= _currentMappings.Count || columnName == ColumnStandard)
		{
			return string.Empty;
		}

		CellRange standardRange = _currentMappings[rowIndex].StandardValueRange;
		if (standardRange == null)
		{
			return "尚未设置标准值区域，无法判断行对齐。";
		}

		int standardRows = standardRange.EndRow - standardRange.StartRow + 1;
		int currentRows = range.EndRow - range.StartRow + 1;
		return standardRows == currentRows
			? "与标准值区域行数一致。"
			: $"与标准值区域行数不一致：标准值 {standardRows} 行，当前区域 {currentRows} 行。";
	}

	private void StandardMode_CheckedChanged(object sender, EventArgs e)
	{
		RadioButton radioButton = sender as RadioButton;
		if (radioButton == null || !radioButton.Checked || !(radioButton.Tag is int))
		{
			return;
		}
		int num = (int)radioButton.Tag;
		foreach (RadioButton modeButton in radioButton.Parent.Controls.OfType<RadioButton>())
		{
			modeButton.ForeColor = modeButton.Checked ? Color.White : Color.FromArgb(29, 29, 31);
		}
		if (radioButton.Text == "手动")
		{
			_manualStandardRows.Add(num);
			MeasurementRule rule = GetRule(num);
			if (rule != null && (rule.ManualStandardValues == null || rule.ManualStandardValues.Count == 0))
			{
				rule.ManualStandardValues = new List<ManualStandardValue>
				{
					new ManualStandardValue { PointIndex = 1, Value = rule.FixedStandardValue }
				};
			}
		}
		else
		{
			_manualStandardRows.Remove(num);
			MeasurementRule rule = GetRule(num);
			if (rule != null)
			{
				rule.ManualStandardValues = new List<ManualStandardValue>();
				// Do not restore the cached fixed value: after a restart it may be
				// the manual value. Automatic mode must resolve the current sheet.
				rule.FixedStandardValue = null;
				rule.MeasurementLowerBound = null;
				rule.MeasurementUpperBound = null;
			}
		}
		SetStandardInputsEnabled(radioButton.Parent, _manualStandardRows.Contains(num));
		NotifyGenerationStateChanged();
	}

	private void ManualStandardValueCount_ValueChanged(object sender, EventArgs e)
	{
		if (!(sender is NumericUpDown input) || !(input.Tag is int rowIndex))
		{
			return;
		}
		MeasurementRule rule = GetRule(rowIndex);
		if (rule == null)
		{
			return;
		}
		int count = Decimal.ToInt32(input.Value);
		List<ManualStandardValue> values = rule.ManualStandardValues ?? new List<ManualStandardValue>();
		while (values.Count < count)
		{
			values.Add(new ManualStandardValue
			{
				PointIndex = values.Count + 1,
				Value = values.Count == 0 ? rule.FixedStandardValue : null
			});
		}
		if (values.Count > count)
		{
			values.RemoveRange(count, values.Count - count);
		}
		rule.ManualStandardValues = values;
		if (count != 1)
		{
			rule.MeasurementLowerBound = null;
			rule.MeasurementUpperBound = null;
		}
		SyncLegacyManualStandardValue(rule);
		RefreshManualStandardValueEditor(rowIndex);
		NotifyGenerationStateChanged();
	}

	private void RefreshManualStandardValueEditor(int rowIndex)
	{
		MappingCardLayout layout = _mappingCards.Controls
			.Cast<Control>()
			.Select(control => control.Tag as MappingCardLayout)
			.FirstOrDefault(item => item != null && item.RowIndex == rowIndex);
		if (layout == null)
		{
			return;
		}
		if (layout.IsCollapsed)
		{
			return;
		}

		Panel editor = layout.StandardValuePanel.Controls.OfType<Panel>().FirstOrDefault();
		if (editor == null)
		{
			return;
		}
		foreach (Control control in editor.Controls.OfType<Control>().Where(control => control.Tag is ManualStandardInputTag).ToList())
		{
			editor.Controls.Remove(control);
			control.Dispose();
		}
		AddManualStandardValueInputs(editor, rowIndex, editor.Width);

		int standardPanelHeight = CalculateStandardValuePanelHeight(GetRule(rowIndex));
		int fieldTop = 50 + standardPanelHeight;
		int statusTop = fieldTop + layout.FieldGrid.Height + 6;
		layout.StandardValuePanel.Height = standardPanelHeight;
		editor.Height = standardPanelHeight - 30;
		layout.FieldGrid.Top = fieldTop;
		layout.StatusLabel.Top = statusTop;
		layout.Card.Height = fieldTop + layout.FieldGrid.Height + 12;
		layout.Card.PerformLayout();
		_mappingCards.PerformLayout();
	}

	private void ManualStandardPoint_ValueChanged(object sender, EventArgs e)
	{
		if (!(sender is NumericUpDown input) || !(input.Tag is ManualStandardInputTag tag))
		{
			return;
		}
		ManualStandardValue value = GetManualStandardValue(tag, true);
		if (value == null)
		{
			return;
		}
		value.PointIndex = Decimal.ToInt32(input.Value);
		SyncLegacyManualStandardValue(GetRule(tag.RowIndex));
		NotifyGenerationStateChanged();
	}

	private void ManualStandardPointValue_TextChanged(object sender, EventArgs e)
	{
		if (!(sender is TextBox input) || !(input.Tag is ManualStandardInputTag tag))
		{
			return;
		}
		ManualStandardValue value = GetManualStandardValue(tag, true);
		if (value == null)
		{
			return;
		}
		MeasurementRule rule = GetRule(tag.RowIndex);
		double lowerBound;
		double upperBound;
		if (tag.ValueIndex == 0 &&
			(rule.ManualStandardValues?.Count ?? 0) <= 1 &&
			ManualStandardValueRangeParser.TryParse(input.Text, out lowerBound, out upperBound))
		{
			value.Value = (lowerBound + upperBound) / 2d;
			rule.MeasurementLowerBound = lowerBound;
			rule.MeasurementUpperBound = upperBound;
		}
		else
		{
			double parsed;
			value.Value = TryParseManualNumber(input.Text, out parsed) ? parsed : (double?)null;
			rule.MeasurementLowerBound = null;
			rule.MeasurementUpperBound = null;
		}
		SyncLegacyManualStandardValue(rule);
		NotifyGenerationStateChanged();
	}

	private ManualStandardValue GetManualStandardValue(ManualStandardInputTag tag, bool create)
	{
		MeasurementRule rule = GetRule(tag.RowIndex);
		if (rule == null || tag.ValueIndex < 0)
		{
			return null;
		}
		rule.ManualStandardValues = rule.ManualStandardValues ?? new List<ManualStandardValue>();
		while (create && rule.ManualStandardValues.Count <= tag.ValueIndex)
		{
			rule.ManualStandardValues.Add(new ManualStandardValue
			{
				PointIndex = rule.ManualStandardValues.Count + 1,
				Value = rule.ManualStandardValues.Count == 0 ? rule.FixedStandardValue : null
			});
		}
		return tag.ValueIndex < rule.ManualStandardValues.Count ? rule.ManualStandardValues[tag.ValueIndex] : null;
	}

	private static bool TryParseManualNumber(string text, out double value)
	{
		return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
			double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
	}

	private static void SyncLegacyManualStandardValue(MeasurementRule rule)
	{
		ManualStandardValue first = (rule?.ManualStandardValues ?? new List<ManualStandardValue>())
			.Where(item => item != null && item.Value.HasValue)
			.OrderBy(item => item.PointIndex)
			.FirstOrDefault();
		if (rule != null)
		{
			rule.FixedStandardValue = first?.Value;
		}
	}

	private sealed class ManualStandardInputTag
	{
		public ManualStandardInputTag(int rowIndex, int valueIndex)
		{
			RowIndex = rowIndex;
			ValueIndex = valueIndex;
		}

		public int RowIndex { get; }
		public int ValueIndex { get; }
	}

	private static void SetStandardInputsEnabled(Control parent, bool enabled)
	{
		if (parent == null)
		{
			return;
		}
		foreach (Control control in parent.Controls)
		{
			if (control is TextBox || control is NumericUpDown)
			{
				control.Enabled = enabled;
			}
			if (control.HasChildren)
			{
				SetStandardInputsEnabled(control, enabled);
			}
		}
	}

    }
}
