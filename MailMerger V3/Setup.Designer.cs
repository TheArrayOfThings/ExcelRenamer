namespace MailMerger_V3
{
    partial class Setup
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Setup));
            this.instructionsBox = new System.Windows.Forms.TextBox();
            this.inboxesLabel = new System.Windows.Forms.Label();
            this.inboxesBox = new System.Windows.Forms.ComboBox();
            this.addButton = new System.Windows.Forms.Button();
            this.removeButton = new System.Windows.Forms.Button();
            this.signatureLabel = new System.Windows.Forms.Label();
            this.signatureBox = new System.Windows.Forms.RichTextBox();
            this.signatureMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.hyperlinkToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.insertImageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.insertMergeFieldToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.submitButton = new System.Windows.Forms.Button();
            this.signatureMenu.SuspendLayout();
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
            this.instructionsBox.Size = new System.Drawing.Size(529, 161);
            this.instructionsBox.TabIndex = 100;
            this.instructionsBox.Text = resources.GetString("instructionsBox.Text");
            this.instructionsBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // inboxesLabel
            // 
            this.inboxesLabel.AutoSize = true;
            this.inboxesLabel.Location = new System.Drawing.Point(5, 167);
            this.inboxesLabel.Name = "inboxesLabel";
            this.inboxesLabel.Size = new System.Drawing.Size(44, 13);
            this.inboxesLabel.TabIndex = 1;
            this.inboxesLabel.Text = "Inboxes";
            // 
            // inboxesBox
            // 
            this.inboxesBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.inboxesBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.inboxesBox.FormattingEnabled = true;
            this.inboxesBox.Location = new System.Drawing.Point(63, 167);
            this.inboxesBox.Name = "inboxesBox";
            this.inboxesBox.Size = new System.Drawing.Size(454, 21);
            this.inboxesBox.TabIndex = 0;
            // 
            // addButton
            // 
            this.addButton.Location = new System.Drawing.Point(63, 195);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(75, 23);
            this.addButton.TabIndex = 1;
            this.addButton.Text = "Add";
            this.addButton.UseVisualStyleBackColor = true;
            this.addButton.Click += new System.EventHandler(this.addButton_Click);
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(145, 195);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(75, 23);
            this.removeButton.TabIndex = 2;
            this.removeButton.Text = "Remove";
            this.removeButton.UseVisualStyleBackColor = true;
            this.removeButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // signatureLabel
            // 
            this.signatureLabel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.signatureLabel.AutoSize = true;
            this.signatureLabel.Location = new System.Drawing.Point(5, 245);
            this.signatureLabel.Name = "signatureLabel";
            this.signatureLabel.Size = new System.Drawing.Size(52, 13);
            this.signatureLabel.TabIndex = 5;
            this.signatureLabel.Text = "Signature";
            // 
            // signatureBox
            // 
            this.signatureBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.signatureBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.signatureBox.ContextMenuStrip = this.signatureMenu;
            this.signatureBox.Location = new System.Drawing.Point(63, 225);
            this.signatureBox.Name = "signatureBox";
            this.signatureBox.Size = new System.Drawing.Size(454, 151);
            this.signatureBox.TabIndex = 3;
            this.signatureBox.Text = "";
            // 
            // signatureMenu
            // 
            this.signatureMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem,
            this.hyperlinkToolStripMenuItem,
            this.insertImageToolStripMenuItem,
            this.insertMergeFieldToolStripMenuItem});
            this.signatureMenu.Name = "emailMenu";
            this.signatureMenu.Size = new System.Drawing.Size(169, 114);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.copyToolStripMenuItem.Text = "Copy";
            // 
            // pasteToolStripMenuItem
            // 
            this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            this.pasteToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.pasteToolStripMenuItem.Text = "Paste";
            // 
            // hyperlinkToolStripMenuItem
            // 
            this.hyperlinkToolStripMenuItem.Name = "hyperlinkToolStripMenuItem";
            this.hyperlinkToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.hyperlinkToolStripMenuItem.Text = "Hyperlink";
            // 
            // insertImageToolStripMenuItem
            // 
            this.insertImageToolStripMenuItem.Name = "insertImageToolStripMenuItem";
            this.insertImageToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.insertImageToolStripMenuItem.Text = "Insert Image";
            // 
            // insertMergeFieldToolStripMenuItem
            // 
            this.insertMergeFieldToolStripMenuItem.Name = "insertMergeFieldToolStripMenuItem";
            this.insertMergeFieldToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.insertMergeFieldToolStripMenuItem.Text = "Insert Merge Field";
            // 
            // submitButton
            // 
            this.submitButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.submitButton.Location = new System.Drawing.Point(442, 382);
            this.submitButton.Name = "submitButton";
            this.submitButton.Size = new System.Drawing.Size(75, 23);
            this.submitButton.TabIndex = 4;
            this.submitButton.Text = "Submit";
            this.submitButton.UseVisualStyleBackColor = true;
            this.submitButton.Click += new System.EventHandler(this.submitButton_Click);
            // 
            // Setup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(529, 412);
            this.Controls.Add(this.submitButton);
            this.Controls.Add(this.signatureBox);
            this.Controls.Add(this.signatureLabel);
            this.Controls.Add(this.removeButton);
            this.Controls.Add(this.addButton);
            this.Controls.Add(this.inboxesBox);
            this.Controls.Add(this.inboxesLabel);
            this.Controls.Add(this.instructionsBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Setup";
            this.Text = "Setup";
            this.signatureMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox instructionsBox;
        private System.Windows.Forms.Label inboxesLabel;
        private System.Windows.Forms.ComboBox inboxesBox;
        private System.Windows.Forms.Button addButton;
        private System.Windows.Forms.Button removeButton;
        private System.Windows.Forms.Label signatureLabel;
        private System.Windows.Forms.RichTextBox signatureBox;
        private System.Windows.Forms.Button submitButton;
        private System.Windows.Forms.ContextMenuStrip signatureMenu;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pasteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem hyperlinkToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem insertImageToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem insertMergeFieldToolStripMenuItem;
    }
}