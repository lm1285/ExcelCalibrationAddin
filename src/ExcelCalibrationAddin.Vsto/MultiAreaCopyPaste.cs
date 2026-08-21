using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private MultiAreaPositionTemplateStore _multiAreaPositionTemplateStore;
        private MultiAreaCopyPasteDialog _multiAreaCopyPasteDialog;
        private bool _multiAreaCopyPasteDialogRunMode;

        internal void ShowMultiAreaTemplateSave() { ShowMultiAreaDialog(false); }
        internal void RunMultiAreaCopyPasteDirect()
        {
            try
            {
                _multiAreaPositionTemplateStore = _multiAreaPositionTemplateStore ?? new MultiAreaPositionTemplateStore();
                var templates = _multiAreaPositionTemplateStore.List();
                var template = MatchTemplateForActiveSelection(templates);
                var selection = Application?.Selection as Excel.Range;
                var useTemplateAnchor = template == null && selection != null && selection.Areas.Count == 1 && templates.Count == 1;
                if (template == null && useTemplateAnchor) template = templates[0];
                if (template == null) throw new InvalidOperationException("没有匹配到已保存的多区域模板。请先选中与模板相同的多个不相连区域，或仅选中目标起始单元格。");
                var message = RunSavedTemplate(template, useTemplateAnchor);
                TrySetApplicationStatusBar("多区域复制：" + message);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("[VSTO] Direct multi-area paste failed: " + ex);
                MessageBox.Show(ex.Message, "运行多区域粘贴", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        internal void ShowMultiAreaCopyPasteRun() { RunMultiAreaCopyPasteDirect(); }

        internal void ShowMultiAreaTemplateDelete()
        {
                _multiAreaPositionTemplateStore = _multiAreaPositionTemplateStore ?? new MultiAreaPositionTemplateStore();
            using (var dialog = new MultiAreaTemplateDeleteDialog(_multiAreaPositionTemplateStore))
            {
                dialog.ShowDialog(GetExcelMainWindow());
            }
        }

        private void ShowMultiAreaDialog(bool runMode)
        {
            _multiAreaPositionTemplateStore = _multiAreaPositionTemplateStore ?? new MultiAreaPositionTemplateStore();
            if (_multiAreaCopyPasteDialog != null && !_multiAreaCopyPasteDialog.IsDisposed)
            {
                if (_multiAreaCopyPasteDialogRunMode == runMode)
                {
                    _multiAreaCopyPasteDialog.Activate();
                    return;
                }
                _multiAreaCopyPasteDialog.Close();
            }

            _multiAreaCopyPasteDialog = new MultiAreaCopyPasteDialog(
                _multiAreaPositionTemplateStore, runMode, CaptureTemplateFromSelection, RunSavedTemplate,
                message => TrySetApplicationStatusBar("多区域复制：" + message));
            _multiAreaCopyPasteDialogRunMode = runMode;
            _multiAreaCopyPasteDialog.FormClosed += delegate { _multiAreaCopyPasteDialog = null; };
            _multiAreaCopyPasteDialog.Show(GetExcelMainWindow());
        }

        private MultiAreaPositionTemplate CaptureTemplateFromSelection(string name)
        {
            var selection = Application?.Selection as Excel.Range;
            if (selection == null || Application?.ActiveWorkbook == null || Application?.ActiveSheet == null)
            {
                throw new InvalidOperationException("请先在模板 Excel 中选择一个或多个不相连区域。");
            }

            var areas = new List<AbsoluteAreaPosition>();
            foreach (Excel.Range area in selection.Areas)
            {
                areas.Add(new AbsoluteAreaPosition
                {
                    StartRow = area.Row,
                    StartColumn = area.Column,
                    RowCount = area.Rows.Count,
                    ColumnCount = area.Columns.Count
                });
            }

            var template = MultiAreaPositionTemplate.Create(name, areas);
            template.Validate();
            return template;
        }

        private string RunSavedTemplate(MultiAreaPositionTemplate template)
        {
            return RunSavedTemplate(template, false);
        }

        private string RunSavedTemplate(MultiAreaPositionTemplate template, bool useTemplateAnchor)
        {
            if (template == null) throw new InvalidOperationException("请先选择已保存的多区域模板。");
            var selection = Application?.Selection as Excel.Range;
            var targetSheet = Application?.ActiveSheet as Excel.Worksheet;
            if (selection == null || targetSheet == null) throw new InvalidOperationException("请先在当前工作 Excel 中选择待复制粘贴的区域。");

            template.Validate();
            var ranges = new List<Excel.Range>();
            if (useTemplateAnchor)
            {
                ValidateSheetBounds(template.Resolve(selection.Row, selection.Column));
                foreach (var area in template.Resolve(selection.Row, selection.Column))
                {
                    ranges.Add(targetSheet.Range[
                        targetSheet.Cells[area.StartRow, area.StartColumn],
                        targetSheet.Cells[area.StartRow + area.RowCount - 1, area.StartColumn + area.ColumnCount - 1]]);
                }
            }
            else
            {
                foreach (Excel.Range area in selection.Areas) ranges.Add(area);
            }

            var areaCount = 0;
            foreach (var area in ranges)
            {
                var values = area.Value2;
                area.Value2 = values;
                areaCount++;
            }
            return $"匹配成功，已将当前选中的 {areaCount} 个区域转换为值。";
        }

        private static MultiAreaPositionTemplate MatchTemplateForActiveSelection(IReadOnlyList<MultiAreaPositionTemplate> templates)
        {
            var selection = Globals.ThisAddIn.Application?.Selection as Excel.Range;
            var candidates = (templates ?? new List<MultiAreaPositionTemplate>()).Where(item => item != null && item.Areas != null).ToList();
            if (selection != null)
            {
                var selectedAreas = selection.Areas.Cast<Excel.Range>().Select(area => new { Row = area.Row, Column = area.Column, Rows = area.Rows.Count, Columns = area.Columns.Count }).ToList();
                var anchorRow = selectedAreas.Min(item => item.Row);
                var anchorColumn = selectedAreas.Min(item => item.Column);
                var signature = selectedAreas.Select(item => new { Row = item.Row - anchorRow, Column = item.Column - anchorColumn, Rows = item.Rows, Columns = item.Columns }).OrderBy(item => item.Row).ThenBy(item => item.Column).ToList();
                var matched = candidates.FirstOrDefault(template => template.Areas.Count == signature.Count && template.Areas.OrderBy(item => item.RowOffset).ThenBy(item => item.ColumnOffset).Select(item => new { Row = item.RowOffset, Column = item.ColumnOffset, Rows = item.RowCount, Columns = item.ColumnCount }).SequenceEqual(signature));
                if (matched != null) return matched;
            }
            return null;
        }


        private static void ValidateSheetBounds(IEnumerable<AbsoluteAreaPosition> areas)
        {
            foreach (var area in areas ?? Enumerable.Empty<AbsoluteAreaPosition>())
            {
                if (area.StartRow <= 0 || area.StartColumn <= 0 || area.RowCount <= 0 || area.ColumnCount <= 0 ||
                    (long)area.StartRow + area.RowCount - 1 > 1048576 || (long)area.StartColumn + area.ColumnCount - 1 > 16384)
                    throw new InvalidOperationException("区域超出 Excel 工作表边界。");
            }
        }
    }

    internal sealed class MultiAreaCopyPasteDialog : Form
    {
        private readonly MultiAreaPositionTemplateStore _store;
        private readonly bool _runMode;
        private readonly Func<string, MultiAreaPositionTemplate> _captureTemplate;
        private readonly Func<MultiAreaPositionTemplate, string> _runTemplate;
        private readonly Action<string> _status;
        private readonly ComboBox _templates = new ComboBox();
        private readonly TextBox _name = new TextBox();
        private readonly Label _state = new Label();

        public MultiAreaCopyPasteDialog(MultiAreaPositionTemplateStore store, bool runMode, Func<string, MultiAreaPositionTemplate> captureTemplate, Func<MultiAreaPositionTemplate, string> runTemplate, Action<string> status)
        {
            _store = store; _runMode = runMode; _captureTemplate = captureTemplate; _runTemplate = runTemplate; _status = status ?? delegate { };
            Text = runMode ? "运行多区域复制粘贴" : "保存多区域模板";
            Width = 500; Height = runMode ? 220 : 190; StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
            BuildLayout(); RefreshTemplates();
        }

        private void BuildLayout()
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), ColumnCount = 2, RowCount = 4 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 128)); layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); Controls.Add(layout);
            if (_runMode)
            {
                layout.Controls.Add(new Label { Text = "选择已保存模板", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
                _templates.DropDownStyle = ComboBoxStyle.DropDownList; _templates.Dock = DockStyle.Fill; layout.Controls.Add(_templates, 1, 0);
                layout.Controls.Add(CreateButton("匹配并一键粘贴值", Run_Click), 0, 1); layout.SetColumnSpan(layout.GetControlFromPosition(0, 1), 2);
                _state.Text = "请在工作 Excel 中选中匹配区域的起始单元格。";
            }
            else
            {
                layout.Controls.Add(new Label { Text = "模板名称", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
                _name.Dock = DockStyle.Fill; layout.Controls.Add(_name, 1, 0);
                layout.Controls.Add(CreateButton("保存当前不连续选区", Save_Click), 0, 1); layout.SetColumnSpan(layout.GetControlFromPosition(0, 1), 2);
                _state.Text = "请在模板 Excel 中先选择多个不相连区域。";
            }
            _state.AutoEllipsis = true; _state.ForeColor = System.Drawing.Color.DimGray; _state.Dock = DockStyle.Fill;
            layout.Controls.Add(_state, 0, 2); layout.SetColumnSpan(_state, 2);
            var hint = new Label { Text = _runMode ? "匹配当前选区后直接将选中区域转换为值。" : "模板只保存区域位置和尺寸，不保存单元格值。", AutoSize = true, ForeColor = System.Drawing.Color.DimGray, Dock = DockStyle.Fill };
            layout.Controls.Add(hint, 0, 3); layout.SetColumnSpan(hint, 2);
        }

        private static Button CreateButton(string text, EventHandler handler)
        {
            var button = new Button { Text = text, Dock = DockStyle.Fill, Height = 32, AutoSize = true }; button.Click += handler; return button;
        }

        private void RefreshTemplates()
        {
            _templates.Items.Clear(); foreach (var template in _store.List()) _templates.Items.Add(new TemplateItem(template));
            if (_templates.Items.Count > 0) _templates.SelectedIndex = 0;
        }

        private void Save_Click(object sender, EventArgs e)
        {
            RunAction(delegate
            {
                var name = (_name.Text ?? string.Empty).Trim(); if (string.IsNullOrWhiteSpace(name)) throw new InvalidOperationException("请输入模板名称。");
                var saved = _store.Save(_captureTemplate(name)); _state.Text = "已保存模板：" + saved.Name + "，区域 " + saved.Areas.Count + " 个。"; _status("模板已保存：" + saved.Name);
            });
        }

        private void Run_Click(object sender, EventArgs e)
        {
            RunAction(delegate
            {
                var item = _templates.SelectedItem as TemplateItem; if (item == null) throw new InvalidOperationException("请先选择已保存的模板。");
                _state.Text = _runTemplate(item.Template); _status(_state.Text);
            });
        }

        private void RunAction(Action action)
        {
            try { action(); }
            catch (Exception ex) { Trace.WriteLine("[VSTO] Multi-area operation failed: " + ex); _state.Text = ex.Message; MessageBox.Show(ex.Message, "多区域复制粘贴", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        private sealed class TemplateItem
        {
            public TemplateItem(MultiAreaPositionTemplate template) { Template = template; }
            public MultiAreaPositionTemplate Template { get; }
            public override string ToString() { return Template.Name + "（" + Template.Areas.Count + " 个区域）"; }
        }
    }

    internal sealed class MultiAreaTemplateDeleteDialog : Form
    {
        private readonly MultiAreaPositionTemplateStore _store;
        private readonly ComboBox _templates = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };

        public MultiAreaTemplateDeleteDialog(MultiAreaPositionTemplateStore store)
        {
            _store = store;
            Text = "删除多区域模板";
            Width = 420;
            Height = 150;
            StartPosition = FormStartPosition.CenterParent;
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), ColumnCount = 2, RowCount = 2 };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100)); panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.Controls.Add(new Label { Text = "选择模板", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            panel.Controls.Add(_templates, 1, 0);
            var delete = new Button { Text = "删除", Dock = DockStyle.Fill };
            delete.Click += Delete_Click;
            panel.Controls.Add(delete, 0, 1); panel.SetColumnSpan(delete, 2);
            Controls.Add(panel);
            RefreshTemplates();
        }

        private void RefreshTemplates()
        {
            foreach (var template in _store.List()) _templates.Items.Add(new TemplateItem(template));
            if (_templates.Items.Count > 0) _templates.SelectedIndex = 0;
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            var item = _templates.SelectedItem as TemplateItem;
            if (item == null) return;
            if (MessageBox.Show("确认删除模板“" + item.Template.Name + "”？", "删除模板", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            _store.Delete(item.Template.Id);
            Close();
        }

        private sealed class TemplateItem
        {
            public TemplateItem(MultiAreaPositionTemplate template) { Template = template; }
            public MultiAreaPositionTemplate Template { get; }
            public override string ToString() { return Template.Name; }
        }
    }
}
