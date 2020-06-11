using System;
using System.Windows.Forms;

namespace MailMerger_V3
{
    public partial class ModalPrompt : Form
    {
        public string Result { set; get; }
        string title, caption;
        public ModalPrompt(string titlePara, string captionPara)
        {
            InitializeComponent();
            title = titlePara;
            caption = captionPara;
            captionLabel.Text = caption;
            this.Text = title;
            Result = null;
        }

        private void inputBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                Result = inputBox.Text;
                this.Close();
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            Result = inputBox.Text;
            this.Close();
        }
    }
}
