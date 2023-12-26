using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShiftReportApp1
{
    public partial class CustomMessageBox : Form
    {
        private DialogResult dialogResult = DialogResult.None;

        public CustomMessageBox(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, Font font)
        {
            InitializeComponent(text, caption, buttons, icon, font);
        }

        private void InitializeComponent(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, Font font)
        {
            this.SuspendLayout();
            this.StartPosition = FormStartPosition.CenterParent;
            this.Size = new Size(450, 400);
            this.Text = caption;

            Label label = new Label();
            label.AutoSize = true;
            label.Text = text;
            label.Font = font;

            Button okButton = new Button();
            okButton.Text = "OK";
            okButton.DialogResult = DialogResult.OK;
            okButton.Click += (sender, e) =>
            {
                this.dialogResult = DialogResult.OK;
                this.Close();
            };
            okButton.Size = new Size(120, 50);
            okButton.Font = new Font("Arial", 14);

            Button cancelButton = new Button();
            cancelButton.Text = "Отмена";
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Click += (sender, e) =>
            {
                this.dialogResult = DialogResult.Cancel;
                this.Close();
            };
            cancelButton.Size = new Size(120, 50);
            cancelButton.Font = new Font("Arial", 14);

            TableLayoutPanel layoutPanel = new TableLayoutPanel();
            layoutPanel.Dock = DockStyle.Fill;
            layoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layoutPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layoutPanel.Controls.Add(label, 0, 0);
            layoutPanel.Controls.Add(okButton, 0, 1);
            layoutPanel.Controls.Add(cancelButton, 1, 1);

            this.Controls.Add(layoutPanel);
            this.ResumeLayout(false);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, Font font)
        {
            using (CustomMessageBox customMessageBox = new CustomMessageBox(text, caption, buttons, icon, font))
            {
                customMessageBox.ShowDialog();
                return customMessageBox.dialogResult;
            }
        }
    }
}
