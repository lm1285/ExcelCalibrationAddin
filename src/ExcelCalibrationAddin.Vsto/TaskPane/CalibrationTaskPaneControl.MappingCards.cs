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
	private void BindMappings()
	{
		InitializeCollapsedCardState();
		_mappingCards.Controls.Clear();
		_fieldStatusCards.Clear();
		for (int i = 0; i < _currentMappings.Count; i++)
		{
			_mappingCards.Controls.Add(CreateMappingCard(_currentMappings[i], i));
		}
		_emptyLabel.Visible = _currentMappings.Count == 0;
		_mappingCards.Visible = _currentMappings.Count > 0;
		_highlightedRowIndex = -1;
		_highlightedColumnIndex = -1;
		ResizeMappingCards();
	}

	private Control CreateMappingCard(TemplateRegionMapping mapping, int rowIndex)
	{
		// Follow the viewport width. A fixed 340px minimum makes narrow task
		// panes scroll horizontally and hides the second field column.
		int num = ((_mappingCards.Width <= 0) ? 120 : Math.Max(120, _mappingCards.Width - 8));
		bool isCollapsed = _collapsedCalibrationRows.Contains(rowIndex);
		int standardPanelHeight = CalculateStandardValuePanelHeight(GetRule(rowIndex));
		int fieldTop = 50 + standardPanelHeight;
		const int fieldGridHeight = 210;
		int statusTop = fieldTop + fieldGridHeight + 6;
		Panel panel = new Panel
		{
			BackColor = Color.White,
			BorderStyle = BorderStyle.None,
			Margin = new Padding(0, 0, 0, 10),
			Padding = new Padding(12),
			Size = new Size(num, isCollapsed ? 42 : fieldTop + fieldGridHeight + 12)
		};
		panel.Paint += MappingCard_Paint;
		CheckBox checkBox = new CheckBox
		{
			FlatStyle = FlatStyle.Flat,
			Checked = _selectedCalibrationRows.Contains(rowIndex),
			Enabled = !_featuresBlocked,
			Tag = rowIndex
		};
		checkBox.SetBounds(10, 12, 18, 18);
		checkBox.CheckedChanged += CalibrationItemSelection_CheckedChanged;
		Label label = new Label
		{
			AutoEllipsis = true,
			Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold),
			ForeColor = Color.FromArgb(24, 24, 27),
			Text = string.IsNullOrWhiteSpace(mapping.ProjectName) ? "未命名项目" : mapping.ProjectName
		};
		label.SetBounds(32, 10, panel.Width - 112, 24);
		label.Tag = rowIndex;
		label.Click += CalibrationCardHeader_Click;
		checkBox.FlatAppearance.BorderColor = Color.FromArgb(174, 178, 184);
		checkBox.FlatAppearance.CheckedBackColor = Color.FromArgb(232, 247, 238);
		ExcelToggleSwitch enabledToggle = new ExcelToggleSwitch
		{
			Checked = GetRule(rowIndex)?.IsEnabled != false,
			Enabled = !_featuresBlocked,
			Name = "enabledToggle",
			Tag = rowIndex
		};
		enabledToggle.SetBounds(panel.Width - 70, 10, 34, 22);
		enabledToggle.CheckedChanged += CalibrationItemEnabled_CheckedChanged;
		Panel collapseIndicator = new Panel
		{
			Cursor = Cursors.Hand,
			Name = "collapseIndicator",
			Tag = rowIndex
		};
		collapseIndicator.SetBounds(panel.Width - 30, 8, 18, 24);
		collapseIndicator.Paint += delegate(object sender, PaintEventArgs e)
		{
			DrawCollapseChevron(e.Graphics, collapseIndicator.ClientRectangle, isCollapsed);
		};
		collapseIndicator.Click += CalibrationCardHeader_Click;
		Control control = CreateStandardValuePanel(rowIndex, panel.Width - 24);
		control.SetBounds(12, 40, panel.Width - 24, standardPanelHeight);
		TableLayoutPanel tableLayoutPanel = new TableLayoutPanel
		{
			BackColor = panel.BackColor,
			ColumnCount = 2,
			RowCount = 5,
			Margin = Padding.Empty,
			Padding = Padding.Empty
		};
		tableLayoutPanel.SetBounds(12, fieldTop, panel.Width - 24, fieldGridHeight);
		tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
		tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
		for (int i = 0; i < 5; i++)
		{
			tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));
		}
		AddFieldStatus(tableLayoutPanel, rowIndex, 0, 0, "项目区域", "Section", mapping.SectionRange);
		AddFieldStatus(tableLayoutPanel, rowIndex, 1, 0, "设定值", "Setpoint", mapping.SetpointValueRange);
		AddFieldStatus(tableLayoutPanel, rowIndex, 0, 1, "标准值", "Standard", mapping.StandardValueRange);
		AddFieldStatus(tableLayoutPanel, rowIndex, 1, 1, "测量值", "Measurement", mapping.MeasurementValueRange);
		AddFieldStatus(
			tableLayoutPanel,
			rowIndex,
			0,
			2,
			"平均值",
			"Average",
			mapping.AverageValueRange);
		AddFieldStatus(
			tableLayoutPanel,
			rowIndex,
			1,
			2,
			ErrorFormulaClassifier.IsMaximumError(GetRule(rowIndex)) ? "最大误差" : "误差",
			"Error",
			mapping.ErrorValueRange);
		AddFieldStatus(tableLayoutPanel, rowIndex, 0, 3, "技术要求", "Requirement", mapping.TechnicalRequirementRange);
		AddFieldStatus(tableLayoutPanel, rowIndex, 1, 3, "不确定度", "Uncertainty", mapping.UncertaintyRange);
		AddFieldStatus(tableLayoutPanel, rowIndex, 0, 4, "量程", "Range", mapping.RangeValueRange);
		AddFieldStatus(tableLayoutPanel, rowIndex, 1, 4, "结论", "Result", mapping.ResultRange);
		Label mappingStatus = new Label
		{
			AutoEllipsis = true,
			Font = new Font("Microsoft YaHei UI", 8f),
			ForeColor = Color.FromArgb(84, 84, 88),
			Text = BuildRuleStructureStatus(GetRule(rowIndex)),
			Visible = false
		};
		mappingStatus.SetBounds(12, statusTop, panel.Width - 24, 18);
		control.Visible = !isCollapsed;
		tableLayoutPanel.Visible = !isCollapsed;
		panel.Controls.Add(checkBox);
		panel.Controls.Add(label);
		panel.Controls.Add(enabledToggle);
		panel.Controls.Add(collapseIndicator);
		panel.Controls.Add(control);
		panel.Controls.Add(tableLayoutPanel);
		panel.Controls.Add(mappingStatus);
		panel.Tag = new MappingCardLayout
		{
			RowIndex = rowIndex,
			Card = panel,
			StandardValuePanel = control,
			FieldGrid = tableLayoutPanel,
			StatusLabel = mappingStatus,
			IsCollapsed = isCollapsed
		};
		return panel;
	}

	private static string BuildRuleStructureStatus(MeasurementRule rule)
	{
		int totalRows = rule?.RowMappings?.Count ?? 0;
		int completeRows = rule?.RowMappings?.Count(item => item != null && item.IsComplete) ?? 0;
		string rowStatus = totalRows == 0 ? "映射未建立" : $"映射 {completeRows}/{totalRows}";
		string formulaStatus = rule?.ErrorFormula?.HasFormula == true ? "误差公式已定位" : "误差公式未定位";
		string requirementFormulaStatus = rule?.ErrorFormula?.TechnicalRequirementFormulaResolved == true
			? "技术要求已定位"
			: "技术要求未定位";
		return rowStatus + " | " + formulaStatus + " | " + requirementFormulaStatus;
	}

	private Control CreateStandardValuePanel(int rowIndex, int width)
	{
		MeasurementRule rule = GetRule(rowIndex);
		int panelHeight = CalculateStandardValuePanelHeight(rule);
		Panel panel = new Panel
		{
			BackColor = Color.White,
			Size = new Size(width, panelHeight)
		};
		Label label = new Label
		{
			AutoSize = false,
			Font = new Font("Microsoft YaHei UI", 8.5f, FontStyle.Bold),
			ForeColor = Color.FromArgb(110, 110, 115),
			Text = "标准值"
		};
		label.SetBounds(0, 2, 52, 22);
		RadioButton radioButton = new RadioButton
		{
			Appearance = Appearance.Button,
			AutoSize = false,
			Checked = !_manualStandardRows.Contains(rowIndex),
			Enabled = !_featuresBlocked,
			FlatStyle = FlatStyle.Flat,
			BackColor = Color.FromArgb(242, 242, 247),
			Font = new Font("Microsoft YaHei UI", 8f, FontStyle.Bold),
			ForeColor = _manualStandardRows.Contains(rowIndex) ? Color.FromArgb(29, 29, 31) : Color.White,
			Tag = rowIndex,
			Text = "自动",
			TextAlign = ContentAlignment.MiddleCenter
		};
		radioButton.FlatAppearance.BorderColor = Color.FromArgb(33, 115, 70);
		radioButton.FlatAppearance.CheckedBackColor = Color.FromArgb(33, 115, 70);
		radioButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 247, 238);
		radioButton.SetBounds(52, 1, 58, 24);
		radioButton.CheckedChanged += StandardMode_CheckedChanged;
		RadioButton radioButton2 = new RadioButton
		{
			Appearance = Appearance.Button,
			AutoSize = false,
			Checked = _manualStandardRows.Contains(rowIndex),
			Enabled = !_featuresBlocked,
			FlatStyle = FlatStyle.Flat,
			BackColor = Color.FromArgb(242, 242, 247),
			Font = new Font("Microsoft YaHei UI", 8f, FontStyle.Bold),
			ForeColor = _manualStandardRows.Contains(rowIndex) ? Color.White : Color.FromArgb(29, 29, 31),
			Tag = rowIndex,
			Text = "手动",
			TextAlign = ContentAlignment.MiddleCenter
		};
		radioButton2.FlatAppearance.BorderColor = Color.FromArgb(198, 200, 204);
		radioButton2.FlatAppearance.CheckedBackColor = Color.FromArgb(33, 115, 70);
		radioButton2.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 247, 238);
		radioButton2.SetBounds(109, 1, 58, 24);
		radioButton2.CheckedChanged += StandardMode_CheckedChanged;
		Label label2 = new Label
		{
			AutoEllipsis = true,
			BackColor = Color.FromArgb(242, 242, 247),
			Font = new Font("Microsoft YaHei UI", 8f),
			ForeColor = Color.FromArgb(84, 84, 88),
			Text = BuildStandardValueSummary(rule),
			TextAlign = ContentAlignment.MiddleLeft
		};
		label2.SetBounds(0, 28, Math.Max(160, width), 20);
		label2.Visible = false;
		panel.Controls.Add(label);
		panel.Controls.Add(radioButton);
		panel.Controls.Add(radioButton2);
		panel.Controls.Add(label2);
		Label label3 = new Label
		{
			AutoEllipsis = true,
			Font = new Font("Microsoft YaHei UI", 8f),
			ForeColor = Color.FromArgb(110, 110, 115),
			Text = "按测量点输入标准值，数量变化会即时保留已填写内容。"
		};
		label3.SetBounds(0, 52, width, 16);
		label3.Visible = false;
		panel.Controls.Add(label3);
		Panel editor = CreateManualStandardValueEditor(rowIndex, width);
		editor.SetBounds(0, 30, width, panelHeight - 30);
		panel.Controls.Add(editor);
		return panel;
	}

	private void InitializeCollapsedCardState()
	{
		_collapsedCalibrationRows.RemoveWhere(rowIndex => rowIndex < 0 || rowIndex >= _currentMappings.Count);
		if (_currentMappings.Count == 0)
		{
			return;
		}
		if (_hasInitializedCollapsedCards)
		{
			return;
		}

		for (int rowIndex = 1; rowIndex < _currentMappings.Count; rowIndex++)
		{
			_collapsedCalibrationRows.Add(rowIndex);
		}
		_hasInitializedCollapsedCards = true;
	}

	private void CalibrationCardHeader_Click(object sender, EventArgs e)
	{
		Control headerControl = sender as Control;
		if (headerControl == null || !(headerControl.Tag is int))
		{
			return;
		}

		int rowIndex = (int)headerControl.Tag;
		if (_collapsedCalibrationRows.Contains(rowIndex))
		{
			_collapsedCalibrationRows.Remove(rowIndex);
		}
		else
		{
			_collapsedCalibrationRows.Add(rowIndex);
		}
		BindMappings();
		UpdateTemplateLibraryButtons();
	}

	private void MappingCard_Paint(object sender, PaintEventArgs e)
	{
		Panel card = sender as Panel;
		if (card == null)
		{
			return;
		}

		Rectangle border = new Rectangle(0, 0, Math.Max(0, card.Width - 1), Math.Max(0, card.Height - 1));
		using (Pen pen = new Pen(Color.FromArgb(214, 216, 220)))
		{
			e.Graphics.DrawRectangle(pen, border);
		}
		MappingCardLayout layout = card.Tag as MappingCardLayout;
		if (layout == null || layout.IsCollapsed)
		{
			return;
		}
		using (Brush brush = new SolidBrush(Color.FromArgb(33, 115, 70)))
		{
			e.Graphics.FillRectangle(brush, 0, 0, 3, card.Height);
		}
	}

	private static void DrawCollapseChevron(Graphics graphics, Rectangle bounds, bool collapsed)
	{
		graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
		int middleX = bounds.Left + bounds.Width / 2;
		int middleY = bounds.Top + bounds.Height / 2;
		Point[] points = collapsed
			? new[] { new Point(middleX - 4, middleY - 2), new Point(middleX, middleY + 2), new Point(middleX + 4, middleY - 2) }
			: new[] { new Point(middleX - 4, middleY + 2), new Point(middleX, middleY - 2), new Point(middleX + 4, middleY + 2) };
		using (Pen pen = new Pen(Color.FromArgb(48, 48, 54), 1.5f))
		{
			graphics.DrawLines(pen, points);
		}
	}

	private static int CalculateStandardValuePanelHeight(MeasurementRule rule)
	{
		int count = Math.Max(1, rule?.ManualStandardValues?.Count ?? 0);
		return 80 + (count - 1) * 28;
	}

	private Panel CreateManualStandardValueEditor(int rowIndex, int width)
	{
		MeasurementRule rule = GetRule(rowIndex);
		int count = Math.Max(1, rule?.ManualStandardValues?.Count ?? 0);
		Panel editor = new Panel
		{
			BackColor = Color.FromArgb(246, 246, 248),
			BorderStyle = BorderStyle.None,
			Tag = rowIndex
		};
		editor.Controls.Add(CreateManualEditorLabel("数量", 10, 3, 52));
		editor.Controls.Add(CreateManualEditorLabel("位置", 82, 3, 60));
		editor.Controls.Add(CreateManualEditorLabel("标准值", 156, 3, 80));
		NumericUpDown input = new NumericUpDown
		{
			Enabled = (!_featuresBlocked && _manualStandardRows.Contains(rowIndex)),
			Font = new Font("Microsoft YaHei UI", 8.5f),
			Minimum = 1,
			Maximum = 20,
			Value = count,
			Tag = rowIndex
		};
		input.SetBounds(10, 20, 54, 22);
		input.ValueChanged += ManualStandardValueCount_ValueChanged;
		editor.Controls.Add(input);
		editor.Controls.Add(new Label
		{
			Text = "个",
			AutoSize = true,
			ForeColor = Color.FromArgb(110, 110, 115),
			Font = new Font("Microsoft YaHei UI", 8f),
			Location = new Point(66, 23)
		});
		AddManualStandardValueInputs(editor, rowIndex, width);
		return editor;
	}

	private static Label CreateManualEditorLabel(string text, int left, int top, int width)
	{
		return new Label
		{
			Text = text == "新增校准项" ? "+ 新增校准项" : text,
			AutoSize = false,
			ForeColor = Color.FromArgb(110, 110, 115),
			Font = new Font("Microsoft YaHei UI", 8f, FontStyle.Bold),
			TextAlign = ContentAlignment.MiddleLeft,
			Location = new Point(left, top),
			Size = new Size(width, 16)
		};
	}

	private void AddManualStandardValueInputs(Panel panel, int rowIndex, int width)
	{
		MeasurementRule rule = GetRule(rowIndex);
		List<ManualStandardValue> values = rule?.ManualStandardValues ?? new List<ManualStandardValue>();
		if (values.Count == 0 && rule?.FixedStandardValue.HasValue == true)
		{
			values = new List<ManualStandardValue>
			{
				new ManualStandardValue { PointIndex = 1, Value = rule.FixedStandardValue }
			};
		}
		int count = Math.Max(1, values.Count);
		for (int index = 0; index < count; index++)
		{
			ManualStandardValue value = index < values.Count ? values[index] : null;
			int top = 20 + index * 28;
			var tag = new ManualStandardInputTag(rowIndex, index);
			NumericUpDown pointInput = new NumericUpDown
			{
				Enabled = !_featuresBlocked && _manualStandardRows.Contains(rowIndex),
				Minimum = 1,
				Maximum = 999,
				Value = Math.Max(1, value?.PointIndex ?? index + 1),
				Tag = tag
			};
			pointInput.BackColor = Color.White;
			pointInput.BorderStyle = BorderStyle.FixedSingle;
			pointInput.Font = new Font("Microsoft YaHei UI", 8.5f);
			pointInput.SetBounds(82, top, 62, 22);
			pointInput.ValueChanged += ManualStandardPoint_ValueChanged;
			var standardValueInput = new TextBox
			{
				Enabled = !_featuresBlocked && _manualStandardRows.Contains(rowIndex),
				Font = new Font("Microsoft YaHei UI", 9f),
				Tag = tag,
				Text = FormatManualStandardValue(rule, value, index)
			};
			standardValueInput.BackColor = Color.White;
			standardValueInput.BorderStyle = BorderStyle.FixedSingle;
			standardValueInput.SetBounds(156, top, Math.Min(132, Math.Max(96, width - 168)), 22);
			standardValueInput.TextChanged += ManualStandardPointValue_TextChanged;
			panel.Controls.Add(pointInput);
			panel.Controls.Add(standardValueInput);
		}
	}

	private static string FormatManualStandardValue(MeasurementRule rule, ManualStandardValue value, int index)
	{
		if (index == 0 &&
			(rule?.ManualStandardValues?.Count ?? 0) <= 1 &&
			rule?.MeasurementLowerBound.HasValue == true &&
			rule.MeasurementUpperBound.HasValue)
		{
			return $"{FormatNullableNumber(rule.MeasurementLowerBound)}~{FormatNullableNumber(rule.MeasurementUpperBound)}";
		}

		return FormatNullableNumber(value?.Value);
	}

	private static Button CreateActionButton(string text)
	{
		bool isPrimary = text == "保存模板" || text == "新增校准项" || text == "设为当前选区";
		bool isDestructive = text == "删除" || text == "清除区域";
		return new Button
		{
			BackColor = isPrimary ? Color.FromArgb(33, 115, 70) : Color.White,
			FlatStyle = FlatStyle.Flat,
			Font = new Font("Microsoft YaHei UI", 8f, FontStyle.Bold),
			ForeColor = isPrimary ? Color.White : (isDestructive ? Color.FromArgb(196, 50, 44) : Color.FromArgb(48, 48, 54)),
			Text = text,
			UseVisualStyleBackColor = false,
			FlatAppearance =
			{
				BorderColor = isPrimary ? Color.FromArgb(33, 115, 70) : (isDestructive ? Color.FromArgb(236, 154, 149) : Color.FromArgb(198, 200, 204)),
				BorderSize = 1
			}
		};
	}

	private void AddFieldStatus(TableLayoutPanel grid, int rowIndex, int column, int row, string label, string columnName, CellRange range)
	{
		grid.Controls.Add(CreateFieldStatusCard(rowIndex, columnName, label, range), column, row);
	}

	private Control CreateFieldStatusCard(int rowIndex, string columnName, string label, CellRange range)
	{
		var tag = new FieldStatusTag(rowIndex, ColumnIndexFromName(columnName), columnName);
		Panel card = new Panel
		{
			Cursor = (_stateFeaturesBlocked ? Cursors.Default : Cursors.Hand),
			Dock = DockStyle.Fill,
			Margin = new Padding(0, 0, 6, 4),
			Padding = new Padding(8, 3, 4, 2),
			Tag = tag
		};
		Label caption = new Label
		{
			AutoEllipsis = false,
			Dock = DockStyle.Top,
			Font = new Font("Microsoft YaHei UI", 8f),
			Height = 16,
			Name = "fieldCaption",
			Text = label,
			TextAlign = ContentAlignment.MiddleLeft
		};
		Label value = new Label
		{
			AutoEllipsis = false,
			Dock = DockStyle.Fill,
			Font = new Font("Microsoft YaHei UI", 8.25f, FontStyle.Bold),
			Name = "fieldRange",
			Text = range != null ? RangeToDisplay(range) : "未识别",
			TextAlign = ContentAlignment.MiddleLeft
		};
		card.Controls.Add(value);
		card.Controls.Add(caption);
		ApplyFieldStatusCardTheme(card, range, false);
		if (!_stateFeaturesBlocked)
		{
			card.Click += FieldStatusCard_Click;
			caption.Click += delegate(object sender, EventArgs e) { FieldStatusCard_Click(card, e); };
			value.Click += delegate(object sender, EventArgs e) { FieldStatusCard_Click(card, e); };
			if (!_readOnlyMode)
			{
				ContextMenuStrip menu = CreateFieldContextMenu(rowIndex, columnName);
				card.ContextMenuStrip = menu;
				caption.ContextMenuStrip = menu;
				value.ContextMenuStrip = menu;
			}
		}
		_fieldStatusCards.Add(card);
		return card;
	}

	private static void ApplyFieldStatusCardTheme(Control card, CellRange range, bool selected)
	{
		Color background = selected
			? Color.FromArgb(33, 115, 70)
			: (range != null ? Color.FromArgb(232, 247, 238) : Color.FromArgb(247, 248, 249));
		Color captionColor = selected ? Color.White : Color.FromArgb(90, 94, 100);
		Color valueColor = selected ? Color.White : (range != null ? Color.FromArgb(28, 112, 65) : Color.FromArgb(118, 118, 128));
		card.BackColor = background;
		foreach (Control child in card.Controls)
		{
			child.BackColor = background;
			child.ForeColor = child.Name == "fieldRange" ? valueColor : captionColor;
		}
	}

	private ContextMenuStrip CreateFieldContextMenu(int rowIndex, string columnName)
	{
		ContextMenuStrip menu = new ContextMenuStrip();
		ToolStripMenuItem useSelectionItem = new ToolStripMenuItem("设为当前选区");
		ToolStripMenuItem clearItem = new ToolStripMenuItem("清除区域");
		ToolStripMenuItem copyPreviousItem = new ToolStripMenuItem("复制上一项结构");

		useSelectionItem.Click += delegate
		{
			_highlightedRowIndex = rowIndex;
			_highlightedColumnIndex = ColumnIndexFromName(columnName);
			ApplyActiveSelectionToField(rowIndex, columnName, true);
		};
		clearItem.Click += delegate
		{
			_highlightedRowIndex = rowIndex;
			_highlightedColumnIndex = ColumnIndexFromName(columnName);
			SetRangeForField(rowIndex, columnName, null);
			RefreshSelectedFieldStatusCard();
		};
		copyPreviousItem.Click += delegate
		{
			CopyPreviousItemStructure(rowIndex);
		};
		copyPreviousItem.Enabled = rowIndex > 0;

		menu.Items.Add(useSelectionItem);
		menu.Items.Add(clearItem);
		menu.Items.Add(new ToolStripSeparator());
		menu.Items.Add(copyPreviousItem);
		return menu;
	}

    }
}
