using System;
using System.Windows.Forms;

namespace MailMerger_V3
{
    public partial class UserLogin : Form
    {
        public string UserEmail { set; get; }
        public string UserPassword { set; get; }
        public UserLogin()
        {
            InitializeComponent();
        }

        private void submitButton_Click(object sender, EventArgs e)
        {
            SubmitCredentials();
        }

        private void passwordTextbox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SubmitCredentials();
            }
        }

        private void submitButton_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                SubmitCredentials();
            }
        }

        private void SubmitCredentials()
        {
            UserEmail = emailTextbox.Text.Trim();
            UserPassword = passwordTextbox.Text.Trim();
            if (String.IsNullOrEmpty(UserEmail))
            {
                MessageBox.Show(@"Please enter your email address");
                return;
            }
            if (String.IsNullOrEmpty(UserPassword))
            {
                MessageBox.Show(@"Please enter your password");
                return;
            }
            if (UsefulTools.IsValidEmail(UserEmail))
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            } else
            {
                MessageBox.Show(@"The email address you entered is invalid");
            }
        }
    }
}
