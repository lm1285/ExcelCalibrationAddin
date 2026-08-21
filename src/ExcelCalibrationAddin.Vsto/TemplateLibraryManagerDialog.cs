using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Repositories;
using ExcelCalibrationAddin.Host.UseCases;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class TemplateLibraryManagerDialog : Form
    {
        private const int StatusColumnIndex = 4;

        private readonly Func<IReadOnlyList<SavedTemplateInfo>> _loadTemplates;
        private readonly Func<string, TemplateLifecycleStatus, bool> _updateTemplateStatus;
        private readonly Func<string, bool> _deleteTemplate;
        private readonly Func<string, GenerationConfiguration> _loadTemplateGenerationConfiguration;
        private readonly Func<string, GenerationConfiguration, bool> _updateTemplateGenerationConfiguration;
        private readonly Func<int> _uploadPendingTemplates;
        private readonly Func<TemplateSyncRunResult> _syncTemplates;
        private readonly Func<string, TemplateConflictResolutionAction, string, bool> _resolveTemplateConflict;
        private readonly Func<string, bool> _exportDiagnostics;
        private readonly Func<GenerationConfiguration> _loadGlobalGenerationConfiguration;
        private readonly Func<SavedTemplateInfo, bool> _editTemplate;
        private readonly Func<string, IReadOnlyList<SampleDataVersion>> _loadSampleVersions;
        private readonly Func<long, bool> _deleteSampleVersion;
        private readonly DataGridView _templateGrid = new DataGridView();

        private bool _loading;

        public TemplateLibraryManagerDialog(
            Func<IReadOnlyList<SavedTemplateInfo>> loadTemplates,
            Func<string, TemplateLifecycleStatus, bool> updateTemplateStatus,
            Func<string, bool> deleteTemplate,
            Func<string, GenerationConfiguration> loadTemplateGenerationConfiguration,
            Func<string, GenerationConfiguration, bool> updateTemplateGenerationConfiguration,
            Func<int> uploadPendingTemplates,
            Func<TemplateSyncRunResult> syncTemplates,
            Func<string, TemplateConflictResolutionAction, string, bool> resolveTemplateConflict,
            Func<string, bool> exportDiagnostics,
            Func<GenerationConfiguration> loadGlobalGenerationConfiguration,
            Func<SavedTemplateInfo, bool> editTemplate,
            Func<string, IReadOnlyList<SampleDataVersion>> loadSampleVersions,
            Func<long, bool> deleteSampleVersion)
        {
            _loadTemplates = loadTemplates ?? throw new ArgumentNullException(nameof(loadTemplates));
            _updateTemplateStatus = updateTemplateStatus ?? throw new ArgumentNullException(nameof(updateTemplateStatus));
            _deleteTemplate = deleteTemplate ?? throw new ArgumentNullException(nameof(deleteTemplate));
            _loadTemplateGenerationConfiguration = loadTemplateGenerationConfiguration ?? throw new ArgumentNullException(nameof(loadTemplateGenerationConfiguration));
            _updateTemplateGenerationConfiguration = updateTemplateGenerationConfiguration ?? throw new ArgumentNullException(nameof(updateTemplateGenerationConfiguration));
            _uploadPendingTemplates = uploadPendingTemplates ?? throw new ArgumentNullException(nameof(uploadPendingTemplates));
            _syncTemplates = syncTemplates ?? throw new ArgumentNullException(nameof(syncTemplates));
            _resolveTemplateConflict = resolveTemplateConflict ?? throw new ArgumentNullException(nameof(resolveTemplateConflict));
            _exportDiagnostics = exportDiagnostics ?? throw new ArgumentNullException(nameof(exportDiagnostics));
            _loadGlobalGenerationConfiguration = loadGlobalGenerationConfiguration ?? throw new ArgumentNullException(nameof(loadGlobalGenerationConfiguration));
            _editTemplate = editTemplate ?? throw new ArgumentNullException(nameof(editTemplate));
            _loadSampleVersions = loadSampleVersions ?? throw new ArgumentNullException(nameof(loadSampleVersions));
            _deleteSampleVersion = deleteSampleVersion ?? throw new ArgumentNullException(nameof(deleteSampleVersion));

            InitializeLayout();
            LoadTemplates();
        }

        private void InitializeLayout()
        {
            Text = "模板库管理";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(1190, 420);
            Font = new Font("Microsoft YaHei UI", 9F);

            _templateGrid.Location = new Point(18, 18);
            _templateGrid.Size = new Size(1154, 330);
            _templateGrid.AllowUserToAddRows = false;
            _templateGrid.AllowUserToDeleteRows = false;
            _templateGrid.AllowUserToResizeRows = false;
            _templateGrid.AutoGenerateColumns = false;
            _templateGrid.BackgroundColor = Color.White;
            _templateGrid.BorderStyle = BorderStyle.FixedSingle;
            _templateGrid.MultiSelect = false;
            _templateGrid.ReadOnly = false;
            _templateGrid.RowHeadersVisible = false;
            _templateGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _templateGrid.EditMode = DataGridViewEditMode.EditOnEnter;
            _templateGrid.CurrentCellDirtyStateChanged += TemplateGrid_CurrentCellDirtyStateChanged;
            _templateGrid.CellValueChanged += TemplateGrid_CellValueChanged;

            _templateGrid.Columns.Add(CreateTextColumn("模板名称", 210, true));
            _templateGrid.Columns.Add(CreateTextColumn("\u6a21\u677f\u540d\u79f0", 155, true));
            _templateGrid.Columns.Add(CreateTextColumn("\u72ec\u7acb\u547d\u540d", 140, true));
            _templateGrid.Columns.Add(CreateTextColumn("\u6a21\u677f\u7f16\u7801", 115, true));
            _templateGrid.Columns.Add(CreateStatusColumn());
            _templateGrid.Columns.Add(CreateTextColumn("同步", 95, true));
            _templateGrid.Columns.Add(CreateTextColumn("规则数", 60, true));
            _templateGrid.Columns.Add(CreateTextColumn("更新时间", 140, true));
            _templateGrid.Columns.Add(CreateTextColumn("指纹", 120, true));

            _templateGrid.Columns[0].HeaderText = "\u6d4b\u91cf\u9886\u57df";

            var syncButton = new Button
            {
                Text = "立即同步",
                Location = new Point(160, 366),
                Size = new Size(100, 30)
            };
            syncButton.Click += SyncButton_Click;

            var conflictButton = new Button
            {
                Text = "处理冲突",
                Location = new Point(268, 366),
                Size = new Size(100, 30)
            };
            conflictButton.Click += ConflictButton_Click;

            var exportButton = new Button
            {
                Text = "导出诊断包",
                Location = new Point(18, 366),
                Size = new Size(132, 30)
            };
            exportButton.Click += ExportButton_Click;

            var deleteButton = new Button
            {
                Text = "删除",
                Location = new Point(378, 366),
                Size = new Size(84, 30)
            };
            deleteButton.Click += DeleteButton_Click;

            var editButton = new Button
            {
                Text = "编辑模板",
                Location = new Point(750, 366),
                Size = new Size(100, 30)
            };
            editButton.Click += EditButton_Click;

            var sampleButton = new Button
            {
                Text = "样本数据",
                Location = new Point(858, 366),
                Size = new Size(100, 30)
            };
            sampleButton.Click += SampleButton_Click;

            var configButton = new Button
            {
                Text = "独立配置",
                Location = new Point(470, 366),
                Size = new Size(92, 30)
            };
            configButton.Click += ConfigureButton_Click;

            var resetConfigButton = new Button
            {
                Text = "使用全局",
                Location = new Point(570, 366),
                Size = new Size(84, 30)
            };
            resetConfigButton.Click += ResetConfigurationButton_Click;

            var closeButton = new Button
            {
                Text = "关闭",
                DialogResult = DialogResult.Cancel,
                Location = new Point(658, 366),
                Size = new Size(84, 30)
            };

            Controls.Add(_templateGrid);
            Controls.Add(exportButton);
            Controls.Add(syncButton);
            Controls.Add(conflictButton);
            Controls.Add(deleteButton);
            Controls.Add(editButton);
            Controls.Add(sampleButton);
            Controls.Add(configButton);
            Controls.Add(resetConfigButton);
            Controls.Add(closeButton);
            CancelButton = closeButton;
        }

        private static DataGridViewTextBoxColumn CreateTextColumn(string headerText, int width, bool readOnly)
        {
            return new DataGridViewTextBoxColumn
            {
                HeaderText = headerText,
                Width = width,
                ReadOnly = readOnly,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private static DataGridViewComboBoxColumn CreateStatusColumn()
        {
            var column = new DataGridViewComboBoxColumn
            {
                HeaderText = "状态",
                Width = 90,
                FlatStyle = FlatStyle.Flat,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            column.Items.Add(TranslateStatus(TemplateLifecycleStatus.Enabled));
            column.Items.Add(TranslateStatus(TemplateLifecycleStatus.Disabled));
            column.Items.Add(TranslateStatus(TemplateLifecycleStatus.Obsolete));
            return column;
        }

        private void LoadTemplates()
        {
            _loading = true;
            try
            {
                _templateGrid.Rows.Clear();
                var templates = _loadTemplates() ?? new List<SavedTemplateInfo>();
                foreach (var template in templates
                    .OrderBy(item => item.DirectoryMetadata?.MeasurementDomain ?? string.Empty)
                    .ThenBy(item => item.DirectoryMetadata?.TemplateName ?? item.TemplateName)
                    .ThenBy(item => item.DirectoryMetadata?.VariantName ?? string.Empty))
                {
                    var displayName = template.TemplateName +
                        (template.HasGenerationConfigurationOverride ? "（独立配置）" : "（全局配置）");
                    var metadata = template.DirectoryMetadata ?? new TemplateDirectoryMetadata();
                    displayName = string.IsNullOrWhiteSpace(metadata.MeasurementDomain)
                        ? "\u672a\u5206\u7c7b"
                        : metadata.MeasurementDomain;
                    var rowIndex = _templateGrid.Rows.Add(
                        displayName,
                        string.IsNullOrWhiteSpace(metadata.TemplateName) ? template.TemplateName : metadata.TemplateName,
                        string.IsNullOrWhiteSpace(metadata.VariantName) ? "\u9ed8\u8ba4\u65b9\u6848" : metadata.VariantName,
                        metadata.TemplateCode ?? string.Empty,
                        TranslateStatus(template.Status),
                        TranslateSyncStatus(template.LocalSyncStatus),
                        template.RuleCount.ToString(),
                        template.UpdatedAt.ToLocalTime().ToString("yyyy-MM-dd HH:mm"),
                        ShortenFingerprint(template.ExactFingerprint));
                    if (!string.IsNullOrWhiteSpace(template.SyncError))
                    {
                        _templateGrid.Rows[rowIndex].Cells[StatusColumnIndex + 1].ToolTipText = template.SyncError;
                    }

                    _templateGrid.Rows[rowIndex].Tag = template;
                }
            }
            finally
            {
                _loading = false;
            }
        }

        private void TemplateGrid_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (_templateGrid.IsCurrentCellDirty && _templateGrid.CurrentCell.ColumnIndex == StatusColumnIndex)
            {
                _templateGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void TemplateGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (_loading || e.RowIndex < 0 || e.ColumnIndex != StatusColumnIndex)
            {
                return;
            }

            var row = _templateGrid.Rows[e.RowIndex];
            var template = row.Tag as SavedTemplateInfo;
            if (template == null)
            {
                return;
            }

            var statusText = Convert.ToString(row.Cells[StatusColumnIndex].Value);
            var status = ParseStatus(statusText);
            if (template.Status == status)
            {
                return;
            }

            if (_updateTemplateStatus(template.ExactFingerprint, status))
            {
                template.Status = status;
                row.Tag = template;
                return;
            }

            MessageBox.Show("模板状态更新失败。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            LoadTemplates();
        }

        private void SyncButton_Click(object sender, EventArgs e)
        {
            try
            {
                var result = _syncTemplates();
                LoadTemplates();
                MessageBox.Show(
                    BuildSyncMessage(result),
                    "模板库同步",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"模板同步失败：{ex.Message}", "模板库同步", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                LoadTemplates();
            }
        }

        private void ConflictButton_Click(object sender, EventArgs e)
        {
            var template = GetSelectedTemplate();
            if (template == null)
            {
                MessageBox.Show("请先选择一个模板。", "模板冲突", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (template.LocalSyncStatus != TemplateSyncStatus.Conflict || !template.HasRemoteConflict)
            {
                MessageBox.Show("当前模板没有可处理的远端冲突。", "模板冲突", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var choice = ShowConflictChoice(template.TemplateName);
            if (!choice.HasValue)
            {
                return;
            }

            var saveAsName = string.Empty;
            if (choice.Value == TemplateConflictResolutionAction.SaveAs &&
                !TryPromptTemplateName(template.TemplateName + "-本地副本", out saveAsName))
            {
                return;
            }

            try
            {
                if (_resolveTemplateConflict(template.ExactFingerprint, choice.Value, saveAsName))
                {
                    LoadTemplates();
                    MessageBox.Show("模板冲突已处理。", "模板冲突", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MessageBox.Show("模板冲突处理失败，可能已被其他操作处理。", "模板冲突", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                LoadTemplates();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"模板冲突处理失败：{ex.Message}", "模板冲突", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                LoadTemplates();
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog
            {
                Filter = "诊断包 (*.zip)|*.zip|所有文件 (*.*)|*.*",
                FileName = $"excel-calibration-diagnostics-{DateTime.Now:yyyyMMddHHmmss}.zip",
                AddExtension = true,
                DefaultExt = "zip"
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    _exportDiagnostics(dialog.FileName);
                    MessageBox.Show("诊断包已导出。", "诊断包", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"诊断包导出失败：{ex.Message}", "诊断包", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private static string BuildSyncMessage(TemplateSyncRunResult result)
        {
            if (result == null)
            {
                return "模板同步未返回结果。";
            }

            if (!result.Succeeded)
            {
                return $"模板同步失败：{result.ErrorMessage}";
            }

            return $"同步完成：更新 {result.AppliedCount} 个，冲突 {result.ConflictCount} 个，忽略 {result.IgnoredCount} 个，失败 {result.FailedCount} 个；待上传成功 {result.PendingUploadsSucceeded} 个。";
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            var template = GetSelectedTemplate();
            if (template == null)
            {
                MessageBox.Show("请先选择一个模板。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show(
                $"确定删除模板“{template.TemplateName}”吗？",
                "模板库管理",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.OK)
            {
                return;
            }

            if (_deleteTemplate(template.ExactFingerprint))
            {
                LoadTemplates();
                return;
            }

            MessageBox.Show("模板不存在或已被删除。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
            LoadTemplates();
        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            var template = GetSelectedTemplate();
            if (template == null)
            {
                MessageBox.Show("请先选择一个模板。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!_editTemplate(template))
            {
                return;
            }
            DialogResult = DialogResult.OK;
            Close();
        }

        private void SampleButton_Click(object sender, EventArgs e)
        {
            var template = GetSelectedTemplate();
            if (template == null)
            {
                MessageBox.Show("请先选择一个模板。", "样本数据", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new SampleDataVersionDialog(
                _loadSampleVersions(template.ExactFingerprint),
                _deleteSampleVersion))
            {
                dialog.ShowDialog(this);
            }
        }

        private void ConfigureButton_Click(object sender, EventArgs e)
        {
            var template = GetSelectedTemplate();
            if (template == null)
            {
                MessageBox.Show("请先选择一个模板。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var configuration = _loadTemplateGenerationConfiguration(template.ExactFingerprint) ??
                _loadGlobalGenerationConfiguration();

            using (var dialog = new RandomGenerationConfigurationDialog(configuration))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                if (_updateTemplateGenerationConfiguration(template.ExactFingerprint, dialog.Configuration))
                {
                    LoadTemplates();
                    return;
                }
            }

            MessageBox.Show("模板随机数配置保存失败。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ResetConfigurationButton_Click(object sender, EventArgs e)
        {
            var template = GetSelectedTemplate();
            if (template == null)
            {
                MessageBox.Show("请先选择一个模板。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!template.HasGenerationConfigurationOverride)
            {
                MessageBox.Show("当前模板已使用全局随机数配置。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show(
                $"确定将模板“{template.TemplateName}”改为使用全局随机数配置吗？",
                "模板库管理",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.OK)
            {
                return;
            }

            if (_updateTemplateGenerationConfiguration(template.ExactFingerprint, null))
            {
                LoadTemplates();
                return;
            }

            MessageBox.Show("模板随机数配置更新失败。", "模板库管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private SavedTemplateInfo GetSelectedTemplate()
        {
            if (_templateGrid.SelectedRows.Count == 0)
            {
                return null;
            }

            return _templateGrid.SelectedRows[0].Tag as SavedTemplateInfo;
        }

        private static TemplateConflictResolutionAction? ShowConflictChoice(string templateName)
        {
            using (var form = new Form())
            using (var label = new Label())
            using (var keepLocalButton = new Button())
            using (var useRemoteButton = new Button())
            using (var saveAsButton = new Button())
            using (var cancelButton = new Button())
            {
                TemplateConflictResolutionAction? selected = null;
                form.Text = "模板冲突";
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.ClientSize = new Size(420, 150);
                form.Font = new Font("Microsoft YaHei UI", 9F);

                label.Text = $"模板“{templateName}”同时存在本地待上传版本和远端更新，请选择处理方式。";
                label.Location = new Point(18, 18);
                label.Size = new Size(384, 48);

                keepLocalButton.Text = "保留本地";
                keepLocalButton.Location = new Point(18, 92);
                keepLocalButton.Size = new Size(88, 30);
                keepLocalButton.Click += (sender, args) =>
                {
                    selected = TemplateConflictResolutionAction.KeepLocal;
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };

                useRemoteButton.Text = "使用远端";
                useRemoteButton.Location = new Point(116, 92);
                useRemoteButton.Size = new Size(88, 30);
                useRemoteButton.Click += (sender, args) =>
                {
                    selected = TemplateConflictResolutionAction.UseRemote;
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };

                saveAsButton.Text = "另存";
                saveAsButton.Location = new Point(214, 92);
                saveAsButton.Size = new Size(78, 30);
                saveAsButton.Click += (sender, args) =>
                {
                    selected = TemplateConflictResolutionAction.SaveAs;
                    form.DialogResult = DialogResult.OK;
                    form.Close();
                };

                cancelButton.Text = "取消";
                cancelButton.Location = new Point(302, 92);
                cancelButton.Size = new Size(78, 30);
                cancelButton.DialogResult = DialogResult.Cancel;

                form.Controls.Add(label);
                form.Controls.Add(keepLocalButton);
                form.Controls.Add(useRemoteButton);
                form.Controls.Add(saveAsButton);
                form.Controls.Add(cancelButton);
                form.CancelButton = cancelButton;

                return form.ShowDialog() == DialogResult.OK ? selected : null;
            }
        }

        private static bool TryPromptTemplateName(string defaultName, out string templateName)
        {
            using (var form = new Form())
            using (var label = new Label())
            using (var textBox = new TextBox())
            using (var okButton = new Button())
            using (var cancelButton = new Button())
            {
                form.Text = "另存模板";
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.ClientSize = new Size(360, 140);
                form.Font = new Font("Microsoft YaHei UI", 9F);

                label.Text = "请输入新模板名称：";
                label.Location = new Point(18, 18);
                label.Size = new Size(320, 24);

                textBox.Text = defaultName ?? string.Empty;
                textBox.Location = new Point(18, 48);
                textBox.Size = new Size(320, 24);

                okButton.Text = "确定";
                okButton.Location = new Point(174, 92);
                okButton.Size = new Size(76, 30);
                okButton.DialogResult = DialogResult.OK;

                cancelButton.Text = "取消";
                cancelButton.Location = new Point(262, 92);
                cancelButton.Size = new Size(76, 30);
                cancelButton.DialogResult = DialogResult.Cancel;

                form.Controls.Add(label);
                form.Controls.Add(textBox);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);
                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                if (form.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(textBox.Text))
                {
                    templateName = string.Empty;
                    return false;
                }

                templateName = textBox.Text.Trim();
                return true;
            }
        }

        private static TemplateLifecycleStatus ParseStatus(string status)
        {
            if (string.Equals(status, "废止", StringComparison.Ordinal))
            {
                return TemplateLifecycleStatus.Obsolete;
            }

            if (string.Equals(status, "停用", StringComparison.Ordinal))
            {
                return TemplateLifecycleStatus.Disabled;
            }

            return TemplateLifecycleStatus.Enabled;
        }

        private static string TranslateStatus(TemplateLifecycleStatus status)
        {
            switch (status)
            {
                case TemplateLifecycleStatus.Disabled:
                    return "停用";
                case TemplateLifecycleStatus.Obsolete:
                    return "废止";
                default:
                    return "启用";
            }
        }

        private static string TranslateSyncStatus(TemplateSyncStatus status)
        {
            switch (status)
            {
                case TemplateSyncStatus.PendingUpload:
                    return "待上传";
                case TemplateSyncStatus.Conflict:
                    return "冲突";
                case TemplateSyncStatus.SyncFailed:
                    return "同步失败";
                default:
                    return "已同步";
            }
        }

        private static string ShortenFingerprint(string fingerprint)
        {
            if (string.IsNullOrWhiteSpace(fingerprint) || fingerprint.Length <= 16)
            {
                return fingerprint ?? string.Empty;
            }

            return fingerprint.Substring(0, 16);
        }
    }
}
