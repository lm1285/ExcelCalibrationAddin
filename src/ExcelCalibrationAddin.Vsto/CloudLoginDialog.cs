using System.Drawing;
using System.Windows.Forms;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class CloudLoginDialog : Form
    {
        private readonly TextBox _username = new TextBox();
        private readonly TextBox _password = new TextBox();

        public CloudLoginDialog()
        {
            Text = "登录器具管理系统";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(360, 174);
            Font = new Font("Microsoft YaHei UI", 9F);

            Controls.Add(new Label { Text = "用户名", Location = new Point(24, 24), AutoSize = true });
            _username.Location = new Point(104, 20);
            _username.Size = new Size(220, 24);
            Controls.Add(_username);

            Controls.Add(new Label { Text = "密码", Location = new Point(24, 64), AutoSize = true });
            _password.Location = new Point(104, 60);
            _password.Size = new Size(220, 24);
            _password.UseSystemPasswordChar = true;
            Controls.Add(_password);

            var ok = new Button { Text = "登录", Location = new Point(164, 116), Size = new Size(76, 30), DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "取消", Location = new Point(248, 116), Size = new Size(76, 30), DialogResult = DialogResult.Cancel };
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        public string Username => _username.Text.Trim();
        public string Password => _password.Text;
    }
}
