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
	private TemplateSaveDirectory PromptTemplateDirectory()
	{
		using (Form form = new Form())
		using (ComboBox domainComboBox = new ComboBox())
		using (Button addDomainButton = new Button())
		using (TextBox templateNameTextBox = new TextBox())
		using (TextBox variantSuffixTextBox = new TextBox())
		using (TextBox codeTextBox = new TextBox())
		using (Button button = new Button())
		using (Button button2 = new Button())
		{
			form.Text = "保存模板";
			form.StartPosition = FormStartPosition.CenterScreen;
			form.FormBorderStyle = FormBorderStyle.FixedDialog;
			form.MaximizeBox = false;
			form.MinimizeBox = false;
			form.ClientSize = new Size(400, 250);
			form.Font = new Font("Microsoft YaHei UI", 9f);
			Label domainLabel = CreateDirectoryLabel("测量领域", 18);
			Label templateLabel = CreateDirectoryLabel("模板名称", 66);
			Label variantLabel = CreateDirectoryLabel("子模板名称", 114);
			Label codeLabel = CreateDirectoryLabel("模板编码", 162);
			domainComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
			domainComboBox.SetBounds(170, 16, 178, 24);
			foreach (string domain in GetSavedMeasurementDomains())
			{
				domainComboBox.Items.Add(domain);
			}
			if (domainComboBox.Items.Count == 0)
			{
				domainComboBox.Items.Add("未分类");
			}
			domainComboBox.SelectedIndex = 0;
			addDomainButton.Text = "+";
			addDomainButton.SetBounds(356, 16, 24, 24);
			addDomainButton.Click += (sender, args) => AddMeasurementDomain(domainComboBox);
			templateNameTextBox.SetBounds(170, 64, 210, 24);
			Label variantPrefixLabel = new Label
			{
				Text = "模板名称 + ",
				Location = new Point(170, 114),
				Size = new Size(110, 22),
				AutoEllipsis = true
			};
			variantSuffixTextBox.Text = "默认方案";
			variantSuffixTextBox.SetBounds(282, 112, 98, 24);
			codeTextBox.SetBounds(170, 160, 210, 24);
			templateNameTextBox.TextChanged += (sender, args) =>
			{
				variantPrefixLabel.Text = templateNameTextBox.Text.Trim() + " + ";
			};
			button.Text = "保存";
			button.DialogResult = DialogResult.OK;
			button.Location = new Point(214, 204);
			button.Size = new Size(76, 28);
			button2.Text = "取消";
			button2.DialogResult = DialogResult.Cancel;
			button2.Location = new Point(302, 204);
			button2.Size = new Size(76, 28);
			form.Controls.Add(domainLabel);
			form.Controls.Add(templateLabel);
			form.Controls.Add(variantLabel);
			form.Controls.Add(codeLabel);
			form.Controls.Add(domainComboBox);
			form.Controls.Add(addDomainButton);
			form.Controls.Add(templateNameTextBox);
			form.Controls.Add(variantPrefixLabel);
			form.Controls.Add(variantSuffixTextBox);
			form.Controls.Add(codeTextBox);
			form.Controls.Add(button);
			form.Controls.Add(button2);
			form.AcceptButton = button;
			form.CancelButton = button2;
			while (form.ShowDialog() == DialogResult.OK)
			{
				string domain = Convert.ToString(domainComboBox.SelectedItem)?.Trim();
				string templateName = templateNameTextBox.Text.Trim();
				string variant = variantSuffixTextBox.Text.Trim();
				if (!string.IsNullOrWhiteSpace(domain) && !string.IsNullOrWhiteSpace(templateName) && !string.IsNullOrWhiteSpace(variant))
				{
					return new TemplateSaveDirectory
					{
						StorageTemplateName = templateName + " + " + variant,
						Metadata = new TemplateDirectoryMetadata
						{
							MeasurementDomain = domain,
							TemplateName = templateName,
							VariantName = variant,
							TemplateCode = codeTextBox.Text.Trim()
						}
					};
				}
				MessageBox.Show("请填写测量领域、模板名称和子模板名称。", "保存模板");
			}
			return null;
		}
	}

	private static IReadOnlyList<string> GetSavedMeasurementDomains()
	{
		try
		{
			return (Globals.ThisAddIn.GetSavedTemplates() ?? new List<SavedTemplateInfo>())
				.Select(item => item.DirectoryMetadata?.MeasurementDomain)
				.Where(item => !string.IsNullOrWhiteSpace(item))
				.Distinct(StringComparer.OrdinalIgnoreCase)
				.OrderBy(item => item)
				.ToList();
		}
		catch
		{
			return new List<string>();
		}
	}

	private static void AddMeasurementDomain(ComboBox domainComboBox)
	{
		string domain = PromptText("新增测量领域", "测量领域");
		if (string.IsNullOrWhiteSpace(domain))
		{
			return;
		}

		int existingIndex = domainComboBox.FindStringExact(domain);
		if (existingIndex < 0)
		{
			domainComboBox.Items.Add(domain);
			existingIndex = domainComboBox.Items.Count - 1;
		}
		domainComboBox.SelectedIndex = existingIndex;
	}

	private static string PromptText(string title, string labelText)
	{
		using (Form form = new Form())
		using (Label label = new Label())
		using (TextBox textBox = new TextBox())
		using (Button confirmButton = new Button())
		using (Button cancelButton = new Button())
		{
			form.Text = title;
			form.StartPosition = FormStartPosition.CenterParent;
			form.FormBorderStyle = FormBorderStyle.FixedDialog;
			form.MaximizeBox = false;
			form.MinimizeBox = false;
			form.ClientSize = new Size(330, 130);
			label.Text = labelText;
			label.SetBounds(18, 18, 294, 22);
			textBox.SetBounds(18, 44, 294, 24);
			confirmButton.Text = "确定";
			confirmButton.DialogResult = DialogResult.OK;
			confirmButton.SetBounds(150, 88, 76, 28);
			cancelButton.Text = "取消";
			cancelButton.DialogResult = DialogResult.Cancel;
			cancelButton.SetBounds(236, 88, 76, 28);
			form.Controls.Add(label);
			form.Controls.Add(textBox);
			form.Controls.Add(confirmButton);
			form.Controls.Add(cancelButton);
			form.AcceptButton = confirmButton;
			form.CancelButton = cancelButton;
			return form.ShowDialog() == DialogResult.OK ? textBox.Text.Trim() : string.Empty;
		}
	}

	private static Label CreateDirectoryLabel(string text, int top)
	{
		return new Label { Text = text, Location = new Point(18, top + 2), Size = new Size(145, 22) };
	}

	private sealed class TemplateSaveDirectory
	{
		public string StorageTemplateName { get; set; }
		public TemplateDirectoryMetadata Metadata { get; set; }
	}

	private bool ConfirmSaveTemplateMode(ref string templateName, out bool createNew)
	{
		createNew = false;
		if (!_canGenerate)
		{
			return true;
		}
		if (!_preferredCreateNewTemplate.HasValue)
		{
			_preferredCreateNewTemplate = PromptExistingTemplateMode();
		}
		if (!_preferredCreateNewTemplate.HasValue)
		{
			return false;
		}
		createNew = _preferredCreateNewTemplate.Value;
		if (createNew)
		{
			templateName = EnsureNewTemplateName(templateName);
		}
		return true;
	}

	private bool? PromptExistingTemplateMode()
	{
		using (Form form = new Form())
		using (Button overwriteButton = new Button())
		using (Button createButton = new Button())
		using (Button cancelButton = new Button())
		{
			form.Text = "保存模板";
			form.StartPosition = FormStartPosition.CenterScreen;
			form.FormBorderStyle = FormBorderStyle.FixedDialog;
			form.MaximizeBox = false;
			form.MinimizeBox = false;
			form.ClientSize = new Size(420, 152);
			form.Font = new Font("Microsoft YaHei UI", 9f);
			Label message = new Label
			{
				Text = "当前模板库已有此模板。请选择覆盖现有模板，或新建一个同指纹模板。",
				Location = new Point(18, 18),
				Size = new Size(384, 48)
			};
			overwriteButton.Text = "覆盖";
			overwriteButton.DialogResult = DialogResult.Yes;
			overwriteButton.Location = new Point(142, 96);
			overwriteButton.Size = new Size(76, 28);
			createButton.Text = "新建";
			createButton.DialogResult = DialogResult.No;
			createButton.Location = new Point(230, 96);
			createButton.Size = new Size(76, 28);
			cancelButton.Text = "取消";
			cancelButton.DialogResult = DialogResult.Cancel;
			cancelButton.Location = new Point(318, 96);
			cancelButton.Size = new Size(76, 28);
			form.Controls.Add(message);
			form.Controls.Add(overwriteButton);
			form.Controls.Add(createButton);
			form.Controls.Add(cancelButton);
			DialogResult result = form.ShowDialog();
			if (result == DialogResult.Cancel)
			{
				return null;
			}
			return result == DialogResult.No;
		}
	}

	private static string EnsureNewTemplateName(string templateName)
	{
		string trimmed = string.IsNullOrWhiteSpace(templateName) ? "新模板" : templateName.Trim();
		return trimmed.EndsWith(" - 新建", StringComparison.Ordinal)
			? trimmed
			: trimmed + " - 新建";
	}

	private MeasurementRule GetRule(int rowIndex)
	{
		return (rowIndex >= 0 && rowIndex < _currentRules.Count) ? _currentRules[rowIndex] : null;
	}

	private void ReindexManualRowsAfterDelete(int deletedRowIndex)
	{
		List<int> list = _manualStandardRows.ToList();
		_manualStandardRows.Clear();
		foreach (int item in list)
		{
			if (item < deletedRowIndex)
			{
				_manualStandardRows.Add(item);
			}
			else if (item > deletedRowIndex)
			{
				_manualStandardRows.Add(item - 1);
			}
		}
		List<KeyValuePair<int, double?>> list2 = _autoStandardValues.ToList();
		_autoStandardValues.Clear();
		foreach (KeyValuePair<int, double?> item2 in list2)
		{
			if (item2.Key < deletedRowIndex)
			{
				_autoStandardValues[item2.Key] = item2.Value;
			}
			else if (item2.Key > deletedRowIndex)
			{
				_autoStandardValues[item2.Key - 1] = item2.Value;
			}
		}
	}

	private void ReindexSelectedRowsAfterDelete(int deletedRowIndex)
	{
		List<int> selectedRows = _selectedCalibrationRows.ToList();
		_selectedCalibrationRows.Clear();
		foreach (int rowIndex in selectedRows)
		{
			if (rowIndex < deletedRowIndex)
			{
				_selectedCalibrationRows.Add(rowIndex);
			}
			else if (rowIndex > deletedRowIndex)
			{
				_selectedCalibrationRows.Add(rowIndex - 1);
			}
		}
	}

    }
}
