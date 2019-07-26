using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailMerger_V3
{
    public partial class SheetSelect : Form
    {
        public string selectedSheet { set; get; }

        public SheetSelect(string[] possibleSheets)
        {
            InitializeComponent();
            foreach (string eachSheet in possibleSheets)
            {
                sheetList.Items.Add(eachSheet);
            }
            sheetList.SelectedItem = sheetList.Items[0];
        }

        private void submitButton_Click(object sender, EventArgs e)
        {
            if (sheetList.SelectedIndex != -1)
            {
                this.DialogResult = DialogResult.OK;
                this.selectedSheet = sheetList.SelectedItem.ToString();
                this.Close();
            } else
            {
                MessageBox.Show("Please select a sheet to import.", "Sheet Selection Needed");
            }

        }
    }
}
