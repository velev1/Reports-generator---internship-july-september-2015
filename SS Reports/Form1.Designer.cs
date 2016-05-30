namespace SS_Reports
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.sourceFilePathBox = new System.Windows.Forms.TextBox();
            this.destinationFileTextBox = new System.Windows.Forms.TextBox();
            this.sourceFileBrowseButton = new System.Windows.Forms.Button();
            this.destinationFileBrowseButton = new System.Windows.Forms.Button();
            this.destinationFileCreateButton = new System.Windows.Forms.Button();
            this.sourceFileClearButton = new System.Windows.Forms.Button();
            this.destinationFileClearButton = new System.Windows.Forms.Button();
            this.sourceFileLabel = new System.Windows.Forms.Label();
            this.destinationFileLabel = new System.Windows.Forms.Label();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.generateReportButton = new System.Windows.Forms.Button();
            this.cancelProcessButton = new System.Windows.Forms.Button();
            this.subtractCheckBox = new System.Windows.Forms.CheckBox();
            this.progressLabel = new System.Windows.Forms.Label();
            this.progressTextBox = new System.Windows.Forms.TextBox();
            this.CreateNewFileWorker = new System.ComponentModel.BackgroundWorker();
            this.browseSourceFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.createNewFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.ReportsManagerWorker = new System.ComponentModel.BackgroundWorker();
            this.openDestinationFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.retailersGroupBox = new System.Windows.Forms.GroupBox();
            this.technomarketRadioButton = new System.Windows.Forms.RadioButton();
            this.technopolisRadioButton = new System.Windows.Forms.RadioButton();
            this.menuStrip.SuspendLayout();
            this.retailersGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // sourceFilePathBox
            // 
            this.sourceFilePathBox.Location = new System.Drawing.Point(94, 29);
            this.sourceFilePathBox.Name = "sourceFilePathBox";
            this.sourceFilePathBox.Size = new System.Drawing.Size(401, 20);
            this.sourceFilePathBox.TabIndex = 0;
            this.sourceFilePathBox.TabStop = false;
            // 
            // destinationFileTextBox
            // 
            this.destinationFileTextBox.Location = new System.Drawing.Point(94, 69);
            this.destinationFileTextBox.Name = "destinationFileTextBox";
            this.destinationFileTextBox.Size = new System.Drawing.Size(401, 20);
            this.destinationFileTextBox.TabIndex = 1;
            this.destinationFileTextBox.TabStop = false;
            // 
            // sourceFileBrowseButton
            // 
            this.sourceFileBrowseButton.Location = new System.Drawing.Point(510, 27);
            this.sourceFileBrowseButton.Name = "sourceFileBrowseButton";
            this.sourceFileBrowseButton.Size = new System.Drawing.Size(75, 23);
            this.sourceFileBrowseButton.TabIndex = 0;
            this.sourceFileBrowseButton.Text = "Browse";
            this.sourceFileBrowseButton.UseVisualStyleBackColor = true;
            this.sourceFileBrowseButton.Click += new System.EventHandler(this.sourceFileBrowseButton_Click);
            // 
            // destinationFileBrowseButton
            // 
            this.destinationFileBrowseButton.Location = new System.Drawing.Point(510, 56);
            this.destinationFileBrowseButton.Name = "destinationFileBrowseButton";
            this.destinationFileBrowseButton.Size = new System.Drawing.Size(75, 23);
            this.destinationFileBrowseButton.TabIndex = 2;
            this.destinationFileBrowseButton.Text = "Browse";
            this.destinationFileBrowseButton.UseVisualStyleBackColor = true;
            this.destinationFileBrowseButton.Click += new System.EventHandler(this.destinationFileBrowseButton_Click);
            // 
            // destinationFileCreateButton
            // 
            this.destinationFileCreateButton.Location = new System.Drawing.Point(510, 79);
            this.destinationFileCreateButton.Name = "destinationFileCreateButton";
            this.destinationFileCreateButton.Size = new System.Drawing.Size(75, 23);
            this.destinationFileCreateButton.TabIndex = 3;
            this.destinationFileCreateButton.Text = "Creater new";
            this.destinationFileCreateButton.UseVisualStyleBackColor = true;
            this.destinationFileCreateButton.Click += new System.EventHandler(this.destinationFileCreateButton_Click);
            // 
            // sourceFileClearButton
            // 
            this.sourceFileClearButton.Location = new System.Drawing.Point(592, 27);
            this.sourceFileClearButton.Name = "sourceFileClearButton";
            this.sourceFileClearButton.Size = new System.Drawing.Size(47, 23);
            this.sourceFileClearButton.TabIndex = 1;
            this.sourceFileClearButton.Text = "Clear";
            this.sourceFileClearButton.UseVisualStyleBackColor = true;
            this.sourceFileClearButton.Click += new System.EventHandler(this.sourceFileClearButton_Click);
            // 
            // destinationFileClearButton
            // 
            this.destinationFileClearButton.Location = new System.Drawing.Point(591, 67);
            this.destinationFileClearButton.Name = "destinationFileClearButton";
            this.destinationFileClearButton.Size = new System.Drawing.Size(46, 23);
            this.destinationFileClearButton.TabIndex = 4;
            this.destinationFileClearButton.Text = "Clear";
            this.destinationFileClearButton.UseVisualStyleBackColor = true;
            this.destinationFileClearButton.Click += new System.EventHandler(this.destinationFileClearButton_Click);
            // 
            // sourceFileLabel
            // 
            this.sourceFileLabel.AutoSize = true;
            this.sourceFileLabel.Location = new System.Drawing.Point(12, 30);
            this.sourceFileLabel.Name = "sourceFileLabel";
            this.sourceFileLabel.Size = new System.Drawing.Size(57, 13);
            this.sourceFileLabel.TabIndex = 7;
            this.sourceFileLabel.Text = "Source file";
            // 
            // destinationFileLabel
            // 
            this.destinationFileLabel.AutoSize = true;
            this.destinationFileLabel.Location = new System.Drawing.Point(12, 70);
            this.destinationFileLabel.Name = "destinationFileLabel";
            this.destinationFileLabel.Size = new System.Drawing.Size(76, 13);
            this.destinationFileLabel.TabIndex = 8;
            this.destinationFileLabel.Text = "Destination file";
            // 
            // menuStrip
            // 
            this.menuStrip.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(659, 24);
            this.menuStrip.TabIndex = 9;
            this.menuStrip.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // generateReportButton
            // 
            this.generateReportButton.Location = new System.Drawing.Point(204, 183);
            this.generateReportButton.Name = "generateReportButton";
            this.generateReportButton.Size = new System.Drawing.Size(90, 34);
            this.generateReportButton.TabIndex = 8;
            this.generateReportButton.Text = "Generate";
            this.generateReportButton.UseVisualStyleBackColor = true;
            this.generateReportButton.Click += new System.EventHandler(this.generateReportButton_Click);
            // 
            // cancelProcessButton
            // 
            this.cancelProcessButton.Location = new System.Drawing.Point(337, 184);
            this.cancelProcessButton.Name = "cancelProcessButton";
            this.cancelProcessButton.Size = new System.Drawing.Size(90, 34);
            this.cancelProcessButton.TabIndex = 9;
            this.cancelProcessButton.Text = "Cancel";
            this.cancelProcessButton.UseVisualStyleBackColor = true;
            this.cancelProcessButton.Click += new System.EventHandler(this.cancelProcessButton_Click);
            // 
            // subtractCheckBox
            // 
            this.subtractCheckBox.AutoSize = true;
            this.subtractCheckBox.Location = new System.Drawing.Point(477, 193);
            this.subtractCheckBox.Name = "subtractCheckBox";
            this.subtractCheckBox.Size = new System.Drawing.Size(66, 17);
            this.subtractCheckBox.TabIndex = 10;
            this.subtractCheckBox.Text = "Subtract";
            this.subtractCheckBox.UseVisualStyleBackColor = true;
            // 
            // progressLabel
            // 
            this.progressLabel.AutoSize = true;
            this.progressLabel.Location = new System.Drawing.Point(28, 228);
            this.progressLabel.Name = "progressLabel";
            this.progressLabel.Size = new System.Drawing.Size(51, 13);
            this.progressLabel.TabIndex = 13;
            this.progressLabel.Text = "Progress:";
            // 
            // progressTextBox
            // 
            this.progressTextBox.BackColor = System.Drawing.Color.White;
            this.progressTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.progressTextBox.Location = new System.Drawing.Point(82, 228);
            this.progressTextBox.Multiline = true;
            this.progressTextBox.Name = "progressTextBox";
            this.progressTextBox.ReadOnly = true;
            this.progressTextBox.Size = new System.Drawing.Size(517, 30);
            this.progressTextBox.TabIndex = 14;
            this.progressTextBox.TabStop = false;
            // 
            // CreateNewFileWorker
            // 
            this.CreateNewFileWorker.WorkerReportsProgress = true;
            this.CreateNewFileWorker.WorkerSupportsCancellation = true;
            this.CreateNewFileWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.NewFileWorker_DoWork);
            this.CreateNewFileWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.NewFileWorker_RunWorkerCompleted);
            // 
            // ReportsManagerWorker
            // 
            this.ReportsManagerWorker.WorkerReportsProgress = true;
            this.ReportsManagerWorker.WorkerSupportsCancellation = true;
            this.ReportsManagerWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.ReportsManagerWorker_DoWork);
            this.ReportsManagerWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.ReportsManagerWorker_RunWorkerCompleted);
            // 
            // retailersGroupBox
            // 
            this.retailersGroupBox.Controls.Add(this.technomarketRadioButton);
            this.retailersGroupBox.Controls.Add(this.technopolisRadioButton);
            this.retailersGroupBox.Location = new System.Drawing.Point(31, 107);
            this.retailersGroupBox.Name = "retailersGroupBox";
            this.retailersGroupBox.Size = new System.Drawing.Size(406, 71);
            this.retailersGroupBox.TabIndex = 5;
            this.retailersGroupBox.TabStop = false;
            this.retailersGroupBox.Text = "Retailers";
            // 
            // technomarketRadioButton
            // 
            this.technomarketRadioButton.AutoSize = true;
            this.technomarketRadioButton.Location = new System.Drawing.Point(17, 42);
            this.technomarketRadioButton.Name = "technomarketRadioButton";
            this.technomarketRadioButton.Size = new System.Drawing.Size(94, 17);
            this.technomarketRadioButton.TabIndex = 7;
            this.technomarketRadioButton.TabStop = true;
            this.technomarketRadioButton.Text = "Technomarket";
            this.technomarketRadioButton.UseVisualStyleBackColor = true;
            // 
            // technopolisRadioButton
            // 
            this.technopolisRadioButton.AutoSize = true;
            this.technopolisRadioButton.Location = new System.Drawing.Point(17, 19);
            this.technopolisRadioButton.Name = "technopolisRadioButton";
            this.technopolisRadioButton.Size = new System.Drawing.Size(83, 17);
            this.technopolisRadioButton.TabIndex = 6;
            this.technopolisRadioButton.TabStop = true;
            this.technopolisRadioButton.Text = "Technopolis";
            this.technopolisRadioButton.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(659, 270);
            this.Controls.Add(this.retailersGroupBox);
            this.Controls.Add(this.progressTextBox);
            this.Controls.Add(this.progressLabel);
            this.Controls.Add(this.subtractCheckBox);
            this.Controls.Add(this.cancelProcessButton);
            this.Controls.Add(this.generateReportButton);
            this.Controls.Add(this.destinationFileLabel);
            this.Controls.Add(this.sourceFileLabel);
            this.Controls.Add(this.destinationFileClearButton);
            this.Controls.Add(this.sourceFileClearButton);
            this.Controls.Add(this.destinationFileCreateButton);
            this.Controls.Add(this.destinationFileBrowseButton);
            this.Controls.Add(this.sourceFileBrowseButton);
            this.Controls.Add(this.destinationFileTextBox);
            this.Controls.Add(this.sourceFilePathBox);
            this.Controls.Add(this.menuStrip);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip;
            this.Name = "Form1";
            this.Text = "SS Reports";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.retailersGroupBox.ResumeLayout(false);
            this.retailersGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox sourceFilePathBox;
        private System.Windows.Forms.TextBox destinationFileTextBox;
        private System.Windows.Forms.Button sourceFileBrowseButton;
        private System.Windows.Forms.Button destinationFileBrowseButton;
        private System.Windows.Forms.Button destinationFileCreateButton;
        private System.Windows.Forms.Button sourceFileClearButton;
        private System.Windows.Forms.Button destinationFileClearButton;
        private System.Windows.Forms.Label sourceFileLabel;
        private System.Windows.Forms.Label destinationFileLabel;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.Button generateReportButton;
        private System.Windows.Forms.Button cancelProcessButton;
        private System.Windows.Forms.CheckBox subtractCheckBox;
        private System.Windows.Forms.Label progressLabel;
        private System.Windows.Forms.TextBox progressTextBox;
        private System.ComponentModel.BackgroundWorker CreateNewFileWorker;
        private System.Windows.Forms.OpenFileDialog browseSourceFileDialog;
        private System.Windows.Forms.SaveFileDialog createNewFileDialog;
        private System.ComponentModel.BackgroundWorker ReportsManagerWorker;
        private System.Windows.Forms.OpenFileDialog openDestinationFileDialog;
        private System.Windows.Forms.GroupBox retailersGroupBox;
        private System.Windows.Forms.RadioButton technomarketRadioButton;
        private System.Windows.Forms.RadioButton technopolisRadioButton;
    }
}

