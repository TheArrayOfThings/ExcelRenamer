using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace MailMerger_V3
{
    public partial class ModalPrompt : Form
    {
        public string Result { set; get; }
        string title, caption;
        bool isPassword;
        public ModalPrompt(string titlePara, string captionPara, bool isPasswordPara)
        {
            InitializeComponent();
            title = titlePara;
            caption = captionPara;
            isPassword = isPasswordPara;
            captionLabel.Text = caption;
            this.Text = title;
            if (isPassword)
            {
                inputBox.PasswordChar = '*';
            }
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
