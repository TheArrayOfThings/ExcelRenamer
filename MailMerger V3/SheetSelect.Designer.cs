namespace MailMerger_V3
{
    partial class SheetSelect
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SheetSelect));
            this.instructionsBox = new System.Windows.Forms.TextBox();
            this.sheetList = new System.Windows.Forms.ComboBox();
            this.submitButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // instructionsBox
            // 
            this.instructionsBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.instructionsBox.BackColor = System.Drawing.SystemColors.Info;
            this.instructionsBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.instructionsBox.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.instructionsBox.Location = new System.Drawing.Point(0, 0);
            this.instructionsBox.Multiline = true;
            this.instructionsBox.Name = "instructionsBox";
            this.instructionsBox.ReadOnly = true;
            this.instructionsBox.Size = new System.Drawing.Size(644, 54);
            this.instructionsBox.TabIndex = 101;
            this.instructionsBox.Text = "\r\nPlease select a sheet from the below dropdown and press \'Submit\'.\r\n";
            this.instructionsBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // sheetList
            // 
            this.sheetList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.sheetList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.sheetList.FormattingEnabled = true;
            this.sheetList.Location = new System.Drawing.Point(13, 60);
            this.sheetList.Name = "sheetList";
            this.sheetList.Size = new System.Drawing.Size(619, 21);
            this.sheetList.TabIndex = 0;
            // 
            // submitButton
            // 
            this.submitButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.submitButton.Location = new System.Drawing.Point(556, 88);
            this.submitButton.Name = "submitButton";
            this.submitButton.Size = new System.Drawing.Size(75, 23);
            this.submitButton.TabIndex = 1;
            this.submitButton.Text = "Submit";
            this.submitButton.UseVisualStyleBackColor = true;
            this.submitButton.Click += new System.EventHandler(this.submitButton_Click);
            // 
            // SheetSelect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(644, 118);
            this.Controls.Add(this.submitButton);
            this.Controls.Add(this.sheetList);
            this.Controls.Add(this.instructionsBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SheetSelect";
            this.Text = "SheetSelect";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox instructionsBox;
        private System.Windows.Forms.ComboBox sheetList;
        private System.Windows.Forms.Button submitButton;
    }
}