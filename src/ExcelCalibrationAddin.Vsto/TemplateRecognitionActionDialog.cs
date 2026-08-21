using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCalibrationAddin.Vsto
{
    internal enum TemplateRecognitionAction
    {
        View,
        Edit,
        Overwrite,
        SaveAs,
        Close
    }

    internal sealed class TemplateRecognitionActionDialog : Form
    {
        private TemplateRecognitionAction _selectedAction = TemplateRecognitionAction.Close;

        private TemplateRecognitionActionDialog(string templateName)
        {
            Text = "模板匹配成功";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(500, 158);
            Font = new Font("Microsoft YaHei UI", 9F);

            var label = new Label
            {
                Text = string.IsNullOrWhiteSpace(templateName)
                    ? "已命中本地模板。请选择后续操作。"
                    : $"已命中模板“{templateName}”。请选择后续操作。",
                Location = new Point(18, 18),
                Size = new Size(464, 44)
            };

            var viewButton = CreateActionButton("查看", 18, TemplateRecognitionAction.View);
            var editButton = CreateActionButton("编辑", 112, TemplateRecognitionAction.Edit);
            var overwriteButton = CreateActionButton("覆盖", 206, TemplateRecognitionAction.Overwrite);
            var saveAsButton = CreateActionButton("另存", 300, TemplateRecognitionAction.SaveAs);
            var closeButton = CreateActionButton("关闭", 394, TemplateRecognitionAction.Close);
            closeButton.DialogResult = DialogResult.Cancel;

            Controls.Add(label);
            Controls.Add(viewButton);
            Controls.Add(editButton);
            Controls.Add(overwriteButton);
            Controls.Add(saveAsButton);
            Controls.Add(closeButton);
            CancelButton = closeButton;
        }

        public static TemplateRecognitionAction ShowDialog(IWin32Window owner, string templateName)
        {
            using (var dialog = new TemplateRecognitionActionDialog(templateName))
            {
                dialog.ShowDialog(owner);
                return dialog._selectedAction;
            }
        }

        private Button CreateActionButton(string text, int left, TemplateRecognitionAction action)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(left, 94),
                Size = new Size(82, 32)
            };
            button.Click += (sender, args) =>
            {
                _selectedAction = action;
                DialogResult = action == TemplateRecognitionAction.Close ? DialogResult.Cancel : DialogResult.OK;
                Close();
            };
            return button;
        }
    }
}
