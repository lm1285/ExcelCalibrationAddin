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
	private void FieldStatusCard_Click(object sender, EventArgs e)
	{
		Control control = sender as Control;
		FieldStatusTag fieldStatusTag = ((control == null) ? null : (control.Tag as FieldStatusTag));
		if (fieldStatusTag != null)
		{
			HighlightMappingCell(fieldStatusTag.RowIndex, fieldStatusTag.ColumnIndex, fieldStatusTag.ColumnName, control);
		}
	}

	private void CalibrationItemSelection_CheckedChanged(object sender, EventArgs e)
	{
		if (!(sender is CheckBox checkBox) || !(checkBox.Tag is int))
		{
			return;
		}
		int rowIndex = (int)checkBox.Tag;
		if (checkBox.Checked)
		{
			_selectedCalibrationRows.Add(rowIndex);
		}
		else
		{
			_selectedCalibrationRows.Remove(rowIndex);
		}
		UpdateTemplateLibraryButtons();
	}

	private void CalibrationItemEnabled_CheckedChanged(object sender, EventArgs e)
	{
		CheckBox toggle = sender as CheckBox;
		if (toggle == null || !(toggle.Tag is int))
		{
			return;
		}

		MeasurementRule rule = GetRule((int)toggle.Tag);
		if (rule == null)
		{
			return;
		}

		toggle.BackColor = toggle.Checked ? Color.FromArgb(33, 115, 70) : Color.White;
		toggle.ForeColor = toggle.Checked ? Color.White : Color.FromArgb(48, 48, 54);
		rule.IsEnabled = toggle.Checked;
		NotifyGenerationStateChanged();
	}

	private void DeleteSelectedCalibrationItemsButton_Click(object sender, EventArgs e)
	{
		if (_featuresBlocked || _selectedCalibrationRows.Count == 0)
		{
			return;
		}
		string message = $"将删除已勾选的 {_selectedCalibrationRows.Count} 个校准项。保存模板后，这些校准项将不再参与后续生成随机数。\r\n\r\n确认删除吗？";
		if (MessageBox.Show(message, "删除校准项", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
		{
			return;
		}
		foreach (int rowIndex in _selectedCalibrationRows.OrderByDescending(item => item).ToList())
		{
			RemoveCalibrationItem(rowIndex);
		}
		_selectedCalibrationRows.Clear();
		BindMappings();
		UpdateTemplateLibraryButtons();
		NotifyGenerationStateChanged();
	}

	private void HighlightMappingCell(int rowIndex, int columnIndex, string columnName, Control selectedStatusCard)
	{
		if (rowIndex >= 0 && rowIndex < _currentMappings.Count && !string.IsNullOrWhiteSpace(columnName))
		{
			MarkFieldStatusSelected(selectedStatusCard);
			CellRange val = ResolveRangeForHighlight(_currentMappings[rowIndex], columnName);
			_highlightedRowIndex = rowIndex;
			_highlightedColumnIndex = columnIndex;
			if (val != null)
			{
				Globals.ThisAddIn.HighlightRange(val);
			}
			UpdateFieldActionBar();
		}
	}

	private void MarkFieldStatusSelected(Control selectedStatusCard)
	{
		foreach (Control fieldStatusCard in _fieldStatusCards)
		{
			FieldStatusTag fieldStatusTag = fieldStatusCard.Tag as FieldStatusTag;
			if (fieldStatusTag != null && fieldStatusTag.RowIndex >= 0 && fieldStatusTag.RowIndex < _currentMappings.Count)
			{
				CellRange val = ResolveRangeForHighlight(_currentMappings[fieldStatusTag.RowIndex], fieldStatusTag.ColumnName);
				ApplyFieldStatusCardTheme(fieldStatusCard, val, false);
			}
		}
		if (selectedStatusCard != null)
		{
			ApplyFieldStatusCardTheme(selectedStatusCard, null, true);
		}
	}

	private void ClearFieldStatusSelection()
	{
		_highlightedRowIndex = -1;
		_highlightedColumnIndex = -1;
		MarkFieldStatusSelected(null);
		UpdateFieldActionBar();
	}

	private void RefreshSelectedFieldStatusCard()
	{
		string columnName = ColumnNameFromIndex(_highlightedColumnIndex);
		Control selectedCard = null;
		foreach (Control fieldStatusCard in _fieldStatusCards)
		{
			FieldStatusTag tag = fieldStatusCard.Tag as FieldStatusTag;
			if (tag == null || tag.RowIndex != _highlightedRowIndex || tag.ColumnName != columnName)
			{
				continue;
			}

			TemplateRegionMapping mapping = _currentMappings[_highlightedRowIndex];
			CellRange range = ResolveRangeForHighlight(mapping, columnName);
			Label caption = fieldStatusCard.Controls.OfType<Label>().FirstOrDefault(control => control.Name == "fieldCaption");
			Label value = fieldStatusCard.Controls.OfType<Label>().FirstOrDefault(control => control.Name == "fieldRange");
			if (caption != null)
			{
				caption.Text = FieldLabelFromColumnName(tag.ColumnName);
			}
			if (value != null)
			{
				value.Text = range != null ? RangeToDisplay(range) : "未识别";
			}
			selectedCard = fieldStatusCard;
			break;
		}

		MarkFieldStatusSelected(selectedCard);
		UpdateTemplateLibraryButtons();
		NotifyGenerationStateChanged();
	}

	private void ResizeMappingCards()
	{
		foreach (Control control3 in _mappingCards.Controls)
		{
			// Cards follow the available pane width. A fixed minimum here creates
			// a horizontal scrollbar on narrow Excel windows and clips the editor.
			control3.Width = ((_mappingCards.ClientSize.Width <= 0) ? control3.Width : Math.Max(120, _mappingCards.ClientSize.Width - 2));
			foreach (Control control4 in control3.Controls)
			{
				if (control4.Name == "collapseIndicator")
				{
					control4.Left = control3.Width - 30;
				}
				else if (control4 is Label)
				{
					control4.Width = Math.Max(0, (control4.Top < 40) ? (control3.Width - 122) : (control3.Width - 24));
				}
				else if (control4.Name == "enabledToggle")
				{
					control4.Left = control3.Width - 70;
				}
				else if (control4 is TableLayoutPanel || control4 is Panel)
				{
					control4.Width = Math.Max(0, control3.Width - 24);
				}
			}
		}
		if (_addCalibrationItemButton != null)
		{
			_addCalibrationItemButton.Left = Math.Max(16, panelRules.Width - 194);
		}
		if (_deleteSelectedCalibrationItemsButton != null)
		{
			_deleteSelectedCalibrationItemsButton.Left = Math.Max(16, panelRules.Width - 86);
		}
	}

    }
}
