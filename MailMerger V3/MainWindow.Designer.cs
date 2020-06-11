namespace MailMerger_V3
{
    partial class main
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(main));
            this.dearLabel = new System.Windows.Forms.Label();
            this.dearBox = new System.Windows.Forms.TextBox();
            this.IDLabel = new System.Windows.Forms.Label();
            this.refBox = new System.Windows.Forms.TextBox();
            this.emailLabel = new System.Windows.Forms.Label();
            this.emailBox = new System.Windows.Forms.TextBox();
            this.previousButton = new System.Windows.Forms.Button();
            this.nextButton = new System.Windows.Forms.Button();
            this.inboxLabel = new System.Windows.Forms.Label();
            this.inboxes = new System.Windows.Forms.ComboBox();
            this.attachmentLabel = new System.Windows.Forms.Label();
            this.attachments = new System.Windows.Forms.ComboBox();
            this.clearButton = new System.Windows.Forms.Button();
            this.addButton = new System.Windows.Forms.Button();
            this.removeButton = new System.Windows.Forms.Button();
            this.subjectBox = new System.Windows.Forms.TextBox();
            this.emailMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.hyperlinkToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.insertImageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.insertMergeFieldToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logBox = new System.Windows.Forms.RichTextBox();
            this.importButton = new System.Windows.Forms.Button();
            this.previewButton = new System.Windows.Forms.Button();
            this.sendButton = new System.Windows.Forms.Button();
            this.settingsButton = new System.Windows.Forms.Button();
            this.bodyBox = new System.Windows.Forms.RichTextBox();
            this.emailMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // dearLabel
            // 
            this.dearLabel.AutoSize = true;
            this.dearLabel.Location = new System.Drawing.Point(16, 11);
            this.dearLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.dearLabel.Name = "dearLabel";
            this.dearLabel.Size = new System.Drawing.Size(39, 17);
            this.dearLabel.TabIndex = 0;
            this.dearLabel.Text = "Dear";
            // 
            // dearBox
            // 
            this.dearBox.Location = new System.Drawing.Point(96, 11);
            this.dearBox.Margin = new System.Windows.Forms.Padding(4);
            this.dearBox.Name = "dearBox";
            this.dearBox.ReadOnly = true;
            this.dearBox.Size = new System.Drawing.Size(299, 22);
            this.dearBox.TabIndex = 1;
            // 
            // IDLabel
            // 
            this.IDLabel.AutoSize = true;
            this.IDLabel.Location = new System.Drawing.Point(16, 43);
            this.IDLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.IDLabel.Name = "IDLabel";
            this.IDLabel.Size = new System.Drawing.Size(74, 17);
            this.IDLabel.TabIndex = 2;
            this.IDLabel.Text = "Reference";
            // 
            // refBox
            // 
            this.refBox.Location = new System.Drawing.Point(96, 43);
            this.refBox.Margin = new System.Windows.Forms.Padding(4);
            this.refBox.Name = "refBox";
            this.refBox.ReadOnly = true;
            this.refBox.Size = new System.Drawing.Size(299, 22);
            this.refBox.TabIndex = 3;
            // 
            // emailLabel
            // 
            this.emailLabel.AutoSize = true;
            this.emailLabel.Location = new System.Drawing.Point(16, 76);
            this.emailLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.emailLabel.Name = "emailLabel";
            this.emailLabel.Size = new System.Drawing.Size(42, 17);
            this.emailLabel.TabIndex = 4;
            this.emailLabel.Text = "Email";
            // 
            // emailBox
            // 
            this.emailBox.Location = new System.Drawing.Point(96, 76);
            this.emailBox.Margin = new System.Windows.Forms.Padding(4);
            this.emailBox.Name = "emailBox";
            this.emailBox.ReadOnly = true;
            this.emailBox.Size = new System.Drawing.Size(299, 22);
            this.emailBox.TabIndex = 5;
            // 
            // previousButton
            // 
            this.previousButton.Location = new System.Drawing.Point(404, 11);
            this.previousButton.Margin = new System.Windows.Forms.Padding(4);
            this.previousButton.Name = "previousButton";
            this.previousButton.Size = new System.Drawing.Size(29, 28);
            this.previousButton.TabIndex = 6;
            this.previousButton.Text = "<";
            this.previousButton.UseVisualStyleBackColor = true;
            this.previousButton.Click += new System.EventHandler(this.PreviousButton_Click);
            // 
            // nextButton
            // 
            this.nextButton.Location = new System.Drawing.Point(443, 11);
            this.nextButton.Margin = new System.Windows.Forms.Padding(4);
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(31, 28);
            this.nextButton.TabIndex = 7;
            this.nextButton.Text = ">";
            this.nextButton.UseVisualStyleBackColor = true;
            this.nextButton.Click += new System.EventHandler(this.NextButton_Click);
            // 
            // inboxLabel
            // 
            this.inboxLabel.AutoSize = true;
            this.inboxLabel.Location = new System.Drawing.Point(493, 11);
            this.inboxLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.inboxLabel.Name = "inboxLabel";
            this.inboxLabel.Size = new System.Drawing.Size(41, 17);
            this.inboxLabel.TabIndex = 8;
            this.inboxLabel.Text = "Inbox";
            // 
            // inboxes
            // 
            this.inboxes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.inboxes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.inboxes.FormattingEnabled = true;
            this.inboxes.Location = new System.Drawing.Point(588, 11);
            this.inboxes.Margin = new System.Windows.Forms.Padding(4);
            this.inboxes.Name = "inboxes";
            this.inboxes.Size = new System.Drawing.Size(457, 24);
            this.inboxes.TabIndex = 9;
            // 
            // attachmentLabel
            // 
            this.attachmentLabel.AutoSize = true;
            this.attachmentLabel.Location = new System.Drawing.Point(493, 43);
            this.attachmentLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.attachmentLabel.Name = "attachmentLabel";
            this.attachmentLabel.Size = new System.Drawing.Size(86, 17);
            this.attachmentLabel.TabIndex = 10;
            this.attachmentLabel.Text = "Attachments";
            // 
            // attachments
            // 
            this.attachments.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.attachments.DisplayMember = "fileName";
            this.attachments.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.attachments.FormattingEnabled = true;
            this.attachments.Location = new System.Drawing.Point(591, 43);
            this.attachments.Margin = new System.Windows.Forms.Padding(4);
            this.attachments.Name = "attachments";
            this.attachments.Size = new System.Drawing.Size(455, 24);
            this.attachments.TabIndex = 11;
            this.attachments.ValueMember = "fileName";
            // 
            // clearButton
            // 
            this.clearButton.Location = new System.Drawing.Point(715, 630);
            this.clearButton.Margin = new System.Windows.Forms.Padding(4);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(69, 28);
            this.clearButton.TabIndex = 12;
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = true;
            this.clearButton.Click += new System.EventHandler(this.ClearButton_Click);
            // 
            // addButton
            // 
            this.addButton.Location = new System.Drawing.Point(591, 75);
            this.addButton.Margin = new System.Windows.Forms.Padding(4);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(100, 28);
            this.addButton.TabIndex = 13;
            this.addButton.Text = "Add";
            this.addButton.UseVisualStyleBackColor = true;
            this.addButton.Click += new System.EventHandler(this.addButton_Click);
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(699, 76);
            this.removeButton.Margin = new System.Windows.Forms.Padding(4);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(100, 28);
            this.removeButton.TabIndex = 14;
            this.removeButton.Text = "Remove";
            this.removeButton.UseVisualStyleBackColor = true;
            this.removeButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // subjectBox
            // 
            this.subjectBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.subjectBox.Location = new System.Drawing.Point(497, 112);
            this.subjectBox.Margin = new System.Windows.Forms.Padding(4);
            this.subjectBox.Name = "subjectBox";
            this.subjectBox.Size = new System.Drawing.Size(548, 22);
            this.subjectBox.TabIndex = 15;
            this.subjectBox.Text = "[Replace with subject]";
            // 
            // emailMenu
            // 
            this.emailMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.emailMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem,
            this.hyperlinkToolStripMenuItem,
            this.insertImageToolStripMenuItem,
            this.insertMergeFieldToolStripMenuItem});
            this.emailMenu.Name = "emailMenu";
            this.emailMenu.Size = new System.Drawing.Size(198, 124);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(197, 24);
            this.copyToolStripMenuItem.Text = "Copy";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyToolStripMenuItem_Click);
            // 
            // pasteToolStripMenuItem
            // 
            this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            this.pasteToolStripMenuItem.Size = new System.Drawing.Size(197, 24);
            this.pasteToolStripMenuItem.Text = "Paste";
            this.pasteToolStripMenuItem.Click += new System.EventHandler(this.pasteToolStripMenuItem_Click);
            // 
            // hyperlinkToolStripMenuItem
            // 
            this.hyperlinkToolStripMenuItem.Name = "hyperlinkToolStripMenuItem";
            this.hyperlinkToolStripMenuItem.Size = new System.Drawing.Size(197, 24);
            this.hyperlinkToolStripMenuItem.Text = "Hyperlink";
            this.hyperlinkToolStripMenuItem.Click += new System.EventHandler(this.hyperlinkToolStripMenuItem_Click);
            // 
            // insertImageToolStripMenuItem
            // 
            this.insertImageToolStripMenuItem.Name = "insertImageToolStripMenuItem";
            this.insertImageToolStripMenuItem.Size = new System.Drawing.Size(197, 24);
            this.insertImageToolStripMenuItem.Text = "Insert Image";
            this.insertImageToolStripMenuItem.Click += new System.EventHandler(this.insertImageToolStripMenuItem_Click);
            // 
            // insertMergeFieldToolStripMenuItem
            // 
            this.insertMergeFieldToolStripMenuItem.Name = "insertMergeFieldToolStripMenuItem";
            this.insertMergeFieldToolStripMenuItem.Size = new System.Drawing.Size(197, 24);
            this.insertMergeFieldToolStripMenuItem.Text = "Insert Merge Field";
            // 
            // logBox
            // 
            this.logBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.logBox.BackColor = System.Drawing.SystemColors.Desktop;
            this.logBox.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logBox.ForeColor = System.Drawing.SystemColors.Window;
            this.logBox.Location = new System.Drawing.Point(20, 112);
            this.logBox.Margin = new System.Windows.Forms.Padding(4);
            this.logBox.Name = "logBox";
            this.logBox.ReadOnly = true;
            this.logBox.Size = new System.Drawing.Size(452, 511);
            this.logBox.TabIndex = 17;
            this.logBox.Text = resources.GetString("logBox.Text");
            // 
            // importButton
            // 
            this.importButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.importButton.Location = new System.Drawing.Point(497, 630);
            this.importButton.Margin = new System.Windows.Forms.Padding(4);
            this.importButton.Name = "importButton";
            this.importButton.Size = new System.Drawing.Size(100, 28);
            this.importButton.TabIndex = 18;
            this.importButton.Text = "Import";
            this.importButton.UseVisualStyleBackColor = true;
            this.importButton.Click += new System.EventHandler(this.ImportButton_Click);
            // 
            // previewButton
            // 
            this.previewButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.previewButton.Location = new System.Drawing.Point(607, 630);
            this.previewButton.Margin = new System.Windows.Forms.Padding(4);
            this.previewButton.Name = "previewButton";
            this.previewButton.Size = new System.Drawing.Size(100, 28);
            this.previewButton.TabIndex = 19;
            this.previewButton.Text = "Preview";
            this.previewButton.UseVisualStyleBackColor = true;
            this.previewButton.Click += new System.EventHandler(this.previewButton_Click);
            // 
            // sendButton
            // 
            this.sendButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.sendButton.Location = new System.Drawing.Point(947, 630);
            this.sendButton.Margin = new System.Windows.Forms.Padding(4);
            this.sendButton.Name = "sendButton";
            this.sendButton.Size = new System.Drawing.Size(100, 28);
            this.sendButton.TabIndex = 20;
            this.sendButton.Text = "Send";
            this.sendButton.UseVisualStyleBackColor = true;
            this.sendButton.Click += new System.EventHandler(this.sendButton_Click);
            // 
            // settingsButton
            // 
            this.settingsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.settingsButton.Location = new System.Drawing.Point(20, 630);
            this.settingsButton.Margin = new System.Windows.Forms.Padding(4);
            this.settingsButton.Name = "settingsButton";
            this.settingsButton.Size = new System.Drawing.Size(100, 28);
            this.settingsButton.TabIndex = 23;
            this.settingsButton.Text = "Settings";
            this.settingsButton.UseVisualStyleBackColor = true;
            this.settingsButton.Click += new System.EventHandler(this.settingsButton_Click);
            // 
            // bodyBox
            // 
            this.bodyBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.bodyBox.ContextMenuStrip = this.emailMenu;
            this.bodyBox.DetectUrls = false;
            this.bodyBox.Location = new System.Drawing.Point(497, 142);
            this.bodyBox.Margin = new System.Windows.Forms.Padding(4);
            this.bodyBox.Name = "bodyBox";
            this.bodyBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth;
            this.bodyBox.Size = new System.Drawing.Size(548, 480);
            this.bodyBox.TabIndex = 16;
            this.bodyBox.Text = "[Replace with body of email]";
            // 
            // main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1053, 677);
            this.Controls.Add(this.settingsButton);
            this.Controls.Add(this.sendButton);
            this.Controls.Add(this.previewButton);
            this.Controls.Add(this.importButton);
            this.Controls.Add(this.logBox);
            this.Controls.Add(this.bodyBox);
            this.Controls.Add(this.subjectBox);
            this.Controls.Add(this.removeButton);
            this.Controls.Add(this.addButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.attachments);
            this.Controls.Add(this.attachmentLabel);
            this.Controls.Add(this.inboxes);
            this.Controls.Add(this.inboxLabel);
            this.Controls.Add(this.nextButton);
            this.Controls.Add(this.previousButton);
            this.Controls.Add(this.emailBox);
            this.Controls.Add(this.emailLabel);
            this.Controls.Add(this.refBox);
            this.Controls.Add(this.IDLabel);
            this.Controls.Add(this.dearBox);
            this.Controls.Add(this.dearLabel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "main";
            this.Text = "Ryan\'s MailMerger V3";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.main_FormClosing);
            this.Load += new System.EventHandler(this.main_Load);
            this.emailMenu.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label dearLabel;
        private System.Windows.Forms.TextBox dearBox;
        private System.Windows.Forms.Label IDLabel;
        private System.Windows.Forms.TextBox refBox;
        private System.Windows.Forms.Label emailLabel;
        private System.Windows.Forms.TextBox emailBox;
        private System.Windows.Forms.Button previousButton;
        private System.Windows.Forms.Button nextButton;
        private System.Windows.Forms.Label inboxLabel;
        private System.Windows.Forms.ComboBox inboxes;
        private System.Windows.Forms.Label attachmentLabel;
        private System.Windows.Forms.ComboBox attachments;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.Button addButton;
        private System.Windows.Forms.Button removeButton;
        private System.Windows.Forms.TextBox subjectBox;
        private System.Windows.Forms.RichTextBox logBox;
        private System.Windows.Forms.Button importButton;
        private System.Windows.Forms.Button previewButton;
        private System.Windows.Forms.Button sendButton;
        private System.Windows.Forms.Button settingsButton;
        private System.Windows.Forms.ContextMenuStrip emailMenu;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pasteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem hyperlinkToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem insertImageToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem insertMergeFieldToolStripMenuItem;
        private System.Windows.Forms.RichTextBox bodyBox;
    }
}

