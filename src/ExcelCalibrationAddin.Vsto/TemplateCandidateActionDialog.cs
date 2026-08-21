using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCalibrationAddin.Vsto
{
    internal enum TemplateCandidateAction
    {
        View,
        ReRecognize,
        SaveAs,
        Close
    }

    internal sealed class TemplateCandidateActionDialog : Form
    {
        private TemplateCandidateAction _selectedAction = TemplateCandidateAction.Close;

        private TemplateCandidateActionDialog(string templateName, double score)
        {
            Text = "\u6a21\u677f\u5019\u9009\u5339\u914d";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(500, 170);
            Font = new Font("Microsoft YaHei UI", 9F);

            var name = string.IsNullOrWhiteSpace(templateName) ? "\u672a\u547d\u540d\u6a21\u677f" : templateName;
            var message = new Label
            {
                Text = string.Format(
                    "\u53d1\u73b0\u76f8\u4f3c\u6a21\u677f\u201c{0}\u201d\uff08\u5339\u914d\u5ea6 {1:F0}%\uff09\u3002\r\n\u5019\u9009\u6a21\u677f\u4e0d\u53ef\u76f4\u63a5\u7528\u4e8e\u751f\u6210\uff0c\u8bf7\u9009\u62e9\u67e5\u770b\u3001\u91cd\u65b0\u8bc6\u522b\u6216\u53e6\u5b58\u4e3a\u65b0\u6a21\u677f\u3002",
                    name,
                    score),
                Location = new Point(18, 18),
                Size = new Size(464, 54)
            };

            Controls.Add(message);
            Controls.Add(CreateActionButton("\u67e5\u770b", 18, TemplateCandidateAction.View));
            Controls.Add(CreateActionButton("\u91cd\u65b0\u8bc6\u522b", 126, TemplateCandidateAction.ReRecognize));
            Controls.Add(CreateActionButton("\u53e6\u5b58", 258, TemplateCandidateAction.SaveAs));
            var closeButton = CreateActionButton("\u5173\u95ed", 390, TemplateCandidateAction.Close);
            closeButton.DialogResult = DialogResult.Cancel;
            Controls.Add(closeButton);
            CancelButton = closeButton;
        }

        public static TemplateCandidateAction ShowDialog(IWin32Window owner, string templateName, double score)
        {
            using (var dialog = new TemplateCandidateActionDialog(templateName, score))
            {
                dialog.ShowDialog(owner);
                return dialog._selectedAction;
            }
        }

        private Button CreateActionButton(string text, int left, TemplateCandidateAction action)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(left, 112),
                Size = new Size(action == TemplateCandidateAction.ReRecognize ? 116 : 82, 32)
            };
            button.Click += (sender, args) =>
            {
                _selectedAction = action;
                DialogResult = action == TemplateCandidateAction.Close ? DialogResult.Cancel : DialogResult.OK;
                Close();
            };
            return button;
        }
    }
}
