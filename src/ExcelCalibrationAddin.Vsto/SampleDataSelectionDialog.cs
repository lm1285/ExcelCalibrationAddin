using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelCalibrationAddin.Contracts;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class SampleDataSelectionDialog : Form
    {
        private readonly CheckedListBox _items = new CheckedListBox();
        public IReadOnlyList<string> SelectedNames => _items.CheckedItems.Cast<string>().ToList();

        public SampleDataSelectionDialog(IEnumerable<MeasurementRule> rules)
        {
            Text = "保存样本数据"; StartPosition = FormStartPosition.CenterParent; ClientSize = new Size(440, 330); Font = new Font("Microsoft YaHei UI", 9F);
            var all = new CheckBox { Text = "全选/取消全选", Location = new Point(18, 16), AutoSize = true };
            all.CheckedChanged += (sender, args) => { for (var i = 0; i < _items.Items.Count; i++) _items.SetItemChecked(i, all.Checked); };
            _items.Location = new Point(18, 48); _items.Size = new Size(404, 220); _items.CheckOnClick = true;
            foreach (var rule in rules ?? Enumerable.Empty<MeasurementRule>()) { var name = string.IsNullOrWhiteSpace(rule?.FieldAlias) ? rule?.FieldName : rule.FieldAlias; if (!string.IsNullOrWhiteSpace(name)) _items.Items.Add(name, true); }
            var ok = new Button { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(252, 282), Size = new Size(80, 30) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(342, 282), Size = new Size(80, 30) };
            Controls.AddRange(new Control[] { all, _items, ok, cancel }); AcceptButton = ok; CancelButton = cancel;
        }
    }

    internal sealed class SampleDataVersionDialog : Form
    {
        private readonly DataGridView _grid = new DataGridView();
        private readonly Func<long, bool> _delete;
        public SampleDataVersionDialog(IEnumerable<ExcelCalibrationAddin.Core.Models.SampleDataVersion> versions, Func<long, bool> delete)
        {
            _delete = delete; Text = "查看样本数据"; StartPosition = FormStartPosition.CenterParent; ClientSize = new Size(680, 360); Font = new Font("Microsoft YaHei UI", 9F);
            _grid.Location = new Point(16, 16); _grid.Size = new Size(648, 280); _grid.ReadOnly = true; _grid.AllowUserToAddRows = false; _grid.RowHeadersVisible = false; _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; _grid.AutoGenerateColumns = false;
            _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "时间", Width = 180 }); _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "校准项数量", Width = 100 }); _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "测量点数量", Width = 100 }); _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "备注", Width = 220 });
            foreach (var version in versions ?? Enumerable.Empty<ExcelCalibrationAddin.Core.Models.SampleDataVersion>()) _grid.Rows.Add(version.CreatedAt.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss"), version.ItemCount, (version.Items ?? new List<ExcelCalibrationAddin.Core.Models.TemplateSampleData>()).Sum(item => item.Points?.Count ?? 0), version.Remark);
            var deleteButton = new Button { Text = "删除选中版本", Location = new Point(16, 310), Size = new Size(120, 30) }; deleteButton.Click += (sender, args) => { var row = _grid.CurrentRow; if (row == null) return; var version = (versions ?? Enumerable.Empty<ExcelCalibrationAddin.Core.Models.SampleDataVersion>()).ElementAtOrDefault(row.Index); if (version != null && MessageBox.Show("确认删除该样本版本？", "查看样本数据", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK && _delete(version.Id)) { _grid.Rows.RemoveAt(row.Index); } };
            var close = new Button { Text = "关闭", DialogResult = DialogResult.Cancel, Location = new Point(568, 310), Size = new Size(96, 30) }; Controls.AddRange(new Control[] { _grid, deleteButton, close }); CancelButton = close;
        }
    }
}
