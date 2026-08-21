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
	private void BuildLayout()
	{
		BackColor = Color.FromArgb(245, 246, 247);
		panelRoot.BackColor = BackColor;
		panelRoot.Padding = Padding.Empty;
		lblHeaderTitle.Text = "校准助手";
		lblHeaderTitle.Font = new Font("Microsoft YaHei UI", 13f, FontStyle.Bold);
		lblHeaderTitle.ForeColor = Color.FromArgb(24, 24, 27);
		lblHeaderSubtitle.Text = "模板识别与规则编辑";
		lblHeaderSubtitle.Font = new Font("Microsoft YaHei UI", 8.5f);
		lblHeaderSubtitle.ForeColor = Color.FromArgb(110, 110, 115);
		lblRemoteCaption.Text = "当前状态";
		lblRemoteCaption.Font = new Font("Microsoft YaHei UI", 8f, FontStyle.Bold);
		lblRemoteCaption.ForeColor = Color.FromArgb(110, 110, 115);
		lblRemoteValue.Font = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold);
		lblRemoteValue.ForeColor = Color.FromArgb(31, 111, 68);
		_matchChainLabel.AutoSize = false;
		_matchChainLabel.Font = new Font("Microsoft YaHei UI", 8f);
		_matchChainLabel.ForeColor = Color.FromArgb(84, 84, 88);
		_statusIcon.AutoSize = false;
		_statusIcon.BackColor = Color.Transparent;
		_statusIcon.Font = new Font("Segoe UI Symbol", 10f, FontStyle.Bold);
		_statusIcon.ForeColor = Color.White;
		_statusIcon.Text = string.Empty;
		_statusIcon.TextAlign = ContentAlignment.MiddleCenter;
		_statusIcon.Paint += StatusIcon_Paint;
		_statusSummary.AutoEllipsis = true;
		_statusSummary.Font = new Font("Microsoft YaHei UI", 8f);
		_statusSummary.ForeColor = Color.FromArgb(110, 110, 115);
		lblRulesTitle.Text = "校准项";
		lblRulesTitle.Font = new Font("Microsoft YaHei UI", 11f, FontStyle.Bold);
		_progressLabel.AutoSize = false;
		_progressLabel.ForeColor = Color.FromArgb(110, 110, 115);
		_progressLabel.Font = new Font("Microsoft YaHei UI", 8.75f);
		_progressLabel.Visible = false;
		_progressBar.Style = ProgressBarStyle.Continuous;
		_progressBar.Height = 10;
		_progressBar.Visible = false;
		_emptyLabel.Text = "尚未识别到校准项";
		_emptyLabel.ForeColor = Color.FromArgb(110, 110, 115);
		_emptyLabel.Font = new Font("Microsoft YaHei UI", 9f);
		_emptyLabel.AutoSize = true;
		ConfigureMappingCards();
		ConfigureTemplateLibraryButtons();
		_addCalibrationItemButton.Text = "+ 新增";
		_overviewCard.BackColor = Color.White;
		_overviewCard.BorderStyle = BorderStyle.None;
		_overviewCard.Paint += OverviewCard_Paint;
		panelOverview.Controls.Add(_overviewCard);
		_overviewCard.Controls.Add(lblRemoteCaption);
		_overviewCard.Controls.Add(lblRemoteValue);
		_overviewCard.Controls.Add(_statusIcon);
		_overviewCard.Controls.Add(_statusSummary);
		_overviewCard.Controls.Add(_progressLabel);
		_overviewCard.Controls.Add(_progressBar);
		_overviewCard.Controls.Add(_saveTemplateButton);
		_overviewCard.Controls.Add(_saveAsTemplateButton);
		_overviewCard.Controls.Add(_matchChainLabel);
		panelRules.Controls.Add(_mappingCards);
		panelRules.Controls.Add(_emptyLabel);
		panelRules.Controls.Add(_addCalibrationItemButton);
		panelRules.Controls.Add(_deleteSelectedCalibrationItemsButton);
		_fieldActionBar.BackColor = Color.White;
		_fieldActionBar.BorderStyle = BorderStyle.None;
		_fieldActionBar.Paint += FieldActionBar_Paint;
		panelRules.Controls.Add(_fieldActionBar);
		_fieldActionBar.Controls.Add(_fieldSelectionLabel);
		_fieldActionBar.Controls.Add(_useSelectionForFieldButton);
		_fieldActionBar.Controls.Add(_clearSelectedFieldRangeButton);
		base.Resize += delegate
		{
			LayoutContent();
		};
		LayoutContent();
	}

	private void ConfigureMappingCards()
	{
		_mappingCards.AutoScroll = true;
		_mappingCards.BackColor = BackColor;
		_mappingCards.BorderStyle = BorderStyle.None;
		_mappingCards.Padding = Padding.Empty;
		_mappingCards.Dock = DockStyle.None;
		_mappingCards.FlowDirection = FlowDirection.TopDown;
		_mappingCards.HorizontalScroll.Enabled = false;
		_mappingCards.HorizontalScroll.Visible = false;
		_mappingCards.WrapContents = false;
	}

	private void ConfigureTemplateLibraryButtons()
	{
		_saveTemplateButton.Click += SaveTemplateButton_Click;
		_saveAsTemplateButton.Click += SaveAsTemplateButton_Click;
		_addCalibrationItemButton.Click += AddCalibrationItemButton_Click;
		_deleteSelectedCalibrationItemsButton.Click += DeleteSelectedCalibrationItemsButton_Click;
		_useSelectionForFieldButton.Click += UseSelectionForField_Click;
		_clearSelectedFieldRangeButton.Click += ClearSelectedFieldRange_Click;
		_fieldSelectionLabel.AutoEllipsis = true;
		_fieldSelectionLabel.Font = new Font("Microsoft YaHei UI", 8f);
		_fieldSelectionLabel.ForeColor = Color.FromArgb(84, 84, 88);
		_fieldSelectionLabel.TextAlign = ContentAlignment.MiddleLeft;
		UpdateTemplateLibraryButtons();
	}

	private void LayoutContent()
	{
		panelHeader.Visible = false;
		panelHeader.Dock = DockStyle.Top;
		panelHeader.Height = 0;
		panelHeader.BackColor = Color.White;
		lblHeaderTitle.Left = 16;
		lblHeaderTitle.Top = 16;
		lblHeaderTitle.Width = base.Width - 32;
		lblHeaderSubtitle.Visible = false;
		panelBody.Dock = DockStyle.Fill;
		panelBody.AutoScroll = true;
		panelBody.Padding = new Padding(0);
		panelBody.BackColor = BackColor;
		panelOverview.Visible = true;
		panelOverview.Dock = DockStyle.Top;
		panelOverview.Height = 84;
		panelOverview.Padding = Padding.Empty;
		panelOverview.BackColor = BackColor;
		_overviewCard.SetBounds(16, 8, Math.Max(120, panelOverview.Width - 32), 68);
		lblRemoteCaption.Visible = false;
		lblRemoteValue.AutoEllipsis = true;
		_statusIcon.SetBounds(14, 17, 24, 24);
		lblRemoteValue.SetBounds(46, 12, Math.Max(80, _overviewCard.Width - 274), 20);
		_statusSummary.SetBounds(46, 36, Math.Max(80, _overviewCard.Width - 274), 18);
		_matchChainLabel.Visible = false;
		_saveTemplateButton.SetBounds(_overviewCard.Width - 214, 18, 94, 28);
		_saveAsTemplateButton.SetBounds(_overviewCard.Width - 112, 18, 98, 28);
		_progressLabel.SetBounds(14, 46, _overviewCard.Width - 28, 16);
		_progressBar.SetBounds(14, 62, _overviewCard.Width - 28, 4);
		panelRules.Visible = true;
		panelRules.Dock = DockStyle.Fill;
		panelRules.MinimumSize = new Size(380, 0);
		panelRules.Padding = new Padding(16);
		panelRules.BackColor = BackColor;
		lblRulesTitle.Visible = true;
		lblRulesTitle.SetBounds(16, 12, Math.Max(0, panelRules.Width - 194), 24);
		_addCalibrationItemButton.Enabled = !_featuresBlocked;
		_deleteSelectedCalibrationItemsButton.Enabled = !_featuresBlocked && _selectedCalibrationRows.Count > 0;
		_addCalibrationItemButton.SetBounds(panelRules.Width - 194, 10, 100, 28);
		_deleteSelectedCalibrationItemsButton.SetBounds(panelRules.Width - 86, 10, 70, 28);
		int actionBarTop = Math.Max(132, panelRules.ClientSize.Height - 54);
		int mappingHeight = Math.Max(76, actionBarTop - 56);
		_mappingCards.SetBounds(16, 48, Math.Max(0, panelRules.Width - 32), mappingHeight);
		_emptyLabel.Location = new Point(18, 58);
		_fieldActionBar.SetBounds(16, actionBarTop, Math.Max(0, panelRules.Width - 32), 42);
		_fieldSelectionLabel.SetBounds(10, 7, Math.Max(80, _fieldActionBar.Width - 222), 28);
		_useSelectionForFieldButton.SetBounds(_fieldActionBar.Width - 202, 7, 112, 28);
		_clearSelectedFieldRangeButton.SetBounds(_fieldActionBar.Width - 80, 7, 70, 28);
		ResizeMappingCards();
	}

	private void OverviewCard_Paint(object sender, PaintEventArgs e)
	{
		Rectangle border = new Rectangle(0, 0, Math.Max(0, _overviewCard.Width - 1), Math.Max(0, _overviewCard.Height - 1));
		using (Pen pen = new Pen(Color.FromArgb(214, 216, 220)))
		{
			e.Graphics.DrawRectangle(pen, border);
		}
	}

	private void StatusIcon_Paint(object sender, PaintEventArgs e)
	{
		Rectangle circle = new Rectangle(1, 1, Math.Max(0, _statusIcon.Width - 2), Math.Max(0, _statusIcon.Height - 2));
		using (Brush brush = new SolidBrush(Color.FromArgb(33, 115, 70)))
		{
			e.Graphics.FillEllipse(brush, circle);
		}
		TextRenderer.DrawText(e.Graphics, "✓", _statusIcon.Font, circle, Color.White, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
	}

	private void FieldActionBar_Paint(object sender, PaintEventArgs e)
	{
		using (Pen pen = new Pen(Color.FromArgb(214, 216, 220)))
		{
			e.Graphics.DrawLine(pen, 0, 0, _fieldActionBar.Width, 0);
		}
	}

    }
}
