namespace S1ExcelPlugIn
{
    partial class FormImportThreats
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormImportThreats));
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.buttonImportAndGenerate = new System.Windows.Forms.Button();
            this.dateTimePickerStart = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickerEnd = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonImportData = new System.Windows.Forms.Button();
            this.checkBoxDateLimit = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxDays = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.numericUpDownMaxRecord = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBoxRecordLimit = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.checkBoxResolved = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownMaxRecord)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(342, 272);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(54, 23);
            this.buttonClose.TabIndex = 2;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // labelStatus
            // 
            this.labelStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelStatus.AutoSize = true;
            this.labelStatus.Location = new System.Drawing.Point(21, 230);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(0, 13);
            this.labelStatus.TabIndex = 6;
            // 
            // buttonImportAndGenerate
            // 
            this.buttonImportAndGenerate.Location = new System.Drawing.Point(125, 272);
            this.buttonImportAndGenerate.Name = "buttonImportAndGenerate";
            this.buttonImportAndGenerate.Size = new System.Drawing.Size(187, 23);
            this.buttonImportAndGenerate.TabIndex = 1;
            this.buttonImportAndGenerate.Tag = "";
            this.buttonImportAndGenerate.Text = "Import Data and Generate Reports";
            this.buttonImportAndGenerate.UseVisualStyleBackColor = true;
            this.buttonImportAndGenerate.Click += new System.EventHandler(this.buttonImportAndGenerate_Click);
            // 
            // dateTimePickerStart
            // 
            this.dateTimePickerStart.Location = new System.Drawing.Point(88, 21);
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.dateTimePickerStart.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerStart.TabIndex = 3;
            this.dateTimePickerStart.ValueChanged += new System.EventHandler(this.dateTimePickerStart_ValueChanged);
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(88, 54);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerEnd.TabIndex = 4;
            this.dateTimePickerEnd.ValueChanged += new System.EventHandler(this.dateTimePickerEnd_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "Start Date:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "End Date:";
            // 
            // buttonImportData
            // 
            this.buttonImportData.Location = new System.Drawing.Point(27, 272);
            this.buttonImportData.Name = "buttonImportData";
            this.buttonImportData.Size = new System.Drawing.Size(75, 23);
            this.buttonImportData.TabIndex = 14;
            this.buttonImportData.Text = "Import Data";
            this.buttonImportData.UseVisualStyleBackColor = true;
            this.buttonImportData.Visible = false;
            this.buttonImportData.Click += new System.EventHandler(this.buttonImportData_Click);
            // 
            // checkBoxDateLimit
            // 
            this.checkBoxDateLimit.AutoSize = true;
            this.checkBoxDateLimit.Location = new System.Drawing.Point(318, 24);
            this.checkBoxDateLimit.Name = "checkBoxDateLimit";
            this.checkBoxDateLimit.Size = new System.Drawing.Size(64, 17);
            this.checkBoxDateLimit.TabIndex = 5;
            this.checkBoxDateLimit.Text = "No Limit";
            this.checkBoxDateLimit.UseVisualStyleBackColor = true;
            this.checkBoxDateLimit.CheckStateChanged += new System.EventHandler(this.checkBoxDateLimit_CheckStateChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxDays);
            this.groupBox1.Controls.Add(this.dateTimePickerStart);
            this.groupBox1.Controls.Add(this.checkBoxDateLimit);
            this.groupBox1.Controls.Add(this.dateTimePickerEnd);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(24, 18);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(445, 93);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Date Range";
            // 
            // textBoxDays
            // 
            this.textBoxDays.Location = new System.Drawing.Point(318, 53);
            this.textBoxDays.Name = "textBoxDays";
            this.textBoxDays.ReadOnly = true;
            this.textBoxDays.Size = new System.Drawing.Size(104, 20);
            this.textBoxDays.TabIndex = 18;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.numericUpDownMaxRecord);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.checkBoxRecordLimit);
            this.groupBox2.Location = new System.Drawing.Point(23, 191);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(445, 62);
            this.groupBox2.TabIndex = 20;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Number of Records";
            // 
            // numericUpDownMaxRecord
            // 
            this.numericUpDownMaxRecord.Location = new System.Drawing.Point(88, 24);
            this.numericUpDownMaxRecord.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numericUpDownMaxRecord.Name = "numericUpDownMaxRecord";
            this.numericUpDownMaxRecord.Size = new System.Drawing.Size(200, 20);
            this.numericUpDownMaxRecord.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 26);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Maximum:";
            // 
            // checkBoxRecordLimit
            // 
            this.checkBoxRecordLimit.AutoSize = true;
            this.checkBoxRecordLimit.Location = new System.Drawing.Point(318, 25);
            this.checkBoxRecordLimit.Name = "checkBoxRecordLimit";
            this.checkBoxRecordLimit.Size = new System.Drawing.Size(64, 17);
            this.checkBoxRecordLimit.TabIndex = 8;
            this.checkBoxRecordLimit.Text = "No Limit";
            this.checkBoxRecordLimit.UseVisualStyleBackColor = true;
            this.checkBoxRecordLimit.CheckedChanged += new System.EventHandler(this.checkBoxRecordLimit_CheckedChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBoxResolved);
            this.groupBox3.Location = new System.Drawing.Point(27, 118);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(441, 57);
            this.groupBox3.TabIndex = 21;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Additional Filters";
            // 
            // checkBoxResolved
            // 
            this.checkBoxResolved.AutoSize = true;
            this.checkBoxResolved.Location = new System.Drawing.Point(91, 24);
            this.checkBoxResolved.Name = "checkBoxResolved";
            this.checkBoxResolved.Size = new System.Drawing.Size(143, 17);
            this.checkBoxResolved.TabIndex = 6;
            this.checkBoxResolved.Text = "Only Unresolved Threats";
            this.checkBoxResolved.UseVisualStyleBackColor = true;
            // 
            // FormImportThreats
            // 
            this.AcceptButton = this.buttonImportAndGenerate;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(493, 314);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonImportData);
            this.Controls.Add(this.buttonImportAndGenerate);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.Name = "FormImportThreats";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Threat Data";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownMaxRecord)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button buttonImportAndGenerate;
        private System.Windows.Forms.DateTimePicker dateTimePickerStart;
        private System.Windows.Forms.DateTimePicker dateTimePickerEnd;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonImportData;
        private System.Windows.Forms.CheckBox checkBoxDateLimit;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBoxDays;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.NumericUpDown numericUpDownMaxRecord;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox checkBoxRecordLimit;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.CheckBox checkBoxResolved;
    }
}