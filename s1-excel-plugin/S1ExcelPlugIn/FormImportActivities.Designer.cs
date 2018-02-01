namespace S1ExcelPlugIn
{
    partial class FormImportActivities
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormImportActivities));
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
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.buttonDelete = new System.Windows.Forms.Button();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.listBoxSelected = new System.Windows.Forms.ListBox();
            this.listBoxTotal = new System.Windows.Forms.ListBox();
            this.buttonUnselectAll = new System.Windows.Forms.Button();
            this.buttonSelectAll = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.numericUpDownMaxRecord = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBoxRecordLimit = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownMaxRecord)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(359, 413);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(54, 23);
            this.buttonClose.TabIndex = 1;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // labelStatus
            // 
            this.labelStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelStatus.AutoSize = true;
            this.labelStatus.Location = new System.Drawing.Point(21, 435);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(0, 13);
            this.labelStatus.TabIndex = 6;
            // 
            // buttonImportAndGenerate
            // 
            this.buttonImportAndGenerate.Location = new System.Drawing.Point(452, 413);
            this.buttonImportAndGenerate.Name = "buttonImportAndGenerate";
            this.buttonImportAndGenerate.Size = new System.Drawing.Size(187, 23);
            this.buttonImportAndGenerate.TabIndex = 13;
            this.buttonImportAndGenerate.Tag = "";
            this.buttonImportAndGenerate.Text = "Import Data and Generate Reports";
            this.buttonImportAndGenerate.UseVisualStyleBackColor = true;
            this.buttonImportAndGenerate.Visible = false;
            this.buttonImportAndGenerate.Click += new System.EventHandler(this.buttonImportAndGenerate_Click);
            // 
            // dateTimePickerStart
            // 
            this.dateTimePickerStart.Location = new System.Drawing.Point(79, 21);
            this.dateTimePickerStart.Name = "dateTimePickerStart";
            this.dateTimePickerStart.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerStart.TabIndex = 8;
            this.dateTimePickerStart.ValueChanged += new System.EventHandler(this.dateTimePickerStart_ValueChanged);
            // 
            // dateTimePickerEnd
            // 
            this.dateTimePickerEnd.Location = new System.Drawing.Point(79, 54);
            this.dateTimePickerEnd.Name = "dateTimePickerEnd";
            this.dateTimePickerEnd.Size = new System.Drawing.Size(200, 20);
            this.dateTimePickerEnd.TabIndex = 9;
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
            this.buttonImportData.Location = new System.Drawing.Point(245, 413);
            this.buttonImportData.Name = "buttonImportData";
            this.buttonImportData.Size = new System.Drawing.Size(75, 23);
            this.buttonImportData.TabIndex = 0;
            this.buttonImportData.Text = "Import Data";
            this.buttonImportData.UseVisualStyleBackColor = true;
            this.buttonImportData.Click += new System.EventHandler(this.buttonImportData_Click);
            // 
            // checkBoxDateLimit
            // 
            this.checkBoxDateLimit.AutoSize = true;
            this.checkBoxDateLimit.Location = new System.Drawing.Point(293, 24);
            this.checkBoxDateLimit.Name = "checkBoxDateLimit";
            this.checkBoxDateLimit.Size = new System.Drawing.Size(64, 17);
            this.checkBoxDateLimit.TabIndex = 10;
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
            this.groupBox1.Location = new System.Drawing.Point(24, 299);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(371, 93);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Date Range";
            // 
            // textBoxDays
            // 
            this.textBoxDays.Location = new System.Drawing.Point(293, 54);
            this.textBoxDays.Name = "textBoxDays";
            this.textBoxDays.ReadOnly = true;
            this.textBoxDays.Size = new System.Drawing.Size(64, 20);
            this.textBoxDays.TabIndex = 18;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.buttonDelete);
            this.groupBox3.Controls.Add(this.buttonAdd);
            this.groupBox3.Controls.Add(this.listBoxSelected);
            this.groupBox3.Controls.Add(this.listBoxTotal);
            this.groupBox3.Controls.Add(this.buttonUnselectAll);
            this.groupBox3.Controls.Add(this.buttonSelectAll);
            this.groupBox3.Location = new System.Drawing.Point(24, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(632, 281);
            this.groupBox3.TabIndex = 20;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Activity Types";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(415, 14);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(118, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Selected Activity Types";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(101, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(114, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Available Activy Types";
            // 
            // buttonDelete
            // 
            this.buttonDelete.Enabled = false;
            this.buttonDelete.Location = new System.Drawing.Point(301, 114);
            this.buttonDelete.Name = "buttonDelete";
            this.buttonDelete.Size = new System.Drawing.Size(28, 23);
            this.buttonDelete.TabIndex = 5;
            this.buttonDelete.Text = "<";
            this.buttonDelete.UseVisualStyleBackColor = true;
            this.buttonDelete.Click += new System.EventHandler(this.buttonDelete_Click);
            // 
            // buttonAdd
            // 
            this.buttonAdd.Enabled = false;
            this.buttonAdd.Location = new System.Drawing.Point(302, 70);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(27, 23);
            this.buttonAdd.TabIndex = 4;
            this.buttonAdd.Text = ">";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // listBoxSelected
            // 
            this.listBoxSelected.DisplayMember = "Text";
            this.listBoxSelected.FormattingEnabled = true;
            this.listBoxSelected.Location = new System.Drawing.Point(335, 32);
            this.listBoxSelected.Name = "listBoxSelected";
            this.listBoxSelected.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxSelected.Size = new System.Drawing.Size(280, 199);
            this.listBoxSelected.Sorted = true;
            this.listBoxSelected.TabIndex = 6;
            this.listBoxSelected.ValueMember = "Value";
            this.listBoxSelected.SelectedIndexChanged += new System.EventHandler(this.listBoxSelected_SelectedIndexChanged);
            // 
            // listBoxTotal
            // 
            this.listBoxTotal.DisplayMember = "Text";
            this.listBoxTotal.FormattingEnabled = true;
            this.listBoxTotal.Location = new System.Drawing.Point(16, 33);
            this.listBoxTotal.Name = "listBoxTotal";
            this.listBoxTotal.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxTotal.Size = new System.Drawing.Size(280, 199);
            this.listBoxTotal.Sorted = true;
            this.listBoxTotal.TabIndex = 2;
            this.listBoxTotal.ValueMember = "Value";
            this.listBoxTotal.SelectedIndexChanged += new System.EventHandler(this.listBoxTotal_SelectedIndexChanged);
            // 
            // buttonUnselectAll
            // 
            this.buttonUnselectAll.Location = new System.Drawing.Point(442, 243);
            this.buttonUnselectAll.Name = "buttonUnselectAll";
            this.buttonUnselectAll.Size = new System.Drawing.Size(75, 23);
            this.buttonUnselectAll.TabIndex = 7;
            this.buttonUnselectAll.Text = "Unselect All";
            this.buttonUnselectAll.UseVisualStyleBackColor = true;
            this.buttonUnselectAll.Click += new System.EventHandler(this.buttonUnselectAll_Click);
            // 
            // buttonSelectAll
            // 
            this.buttonSelectAll.Location = new System.Drawing.Point(122, 243);
            this.buttonSelectAll.Name = "buttonSelectAll";
            this.buttonSelectAll.Size = new System.Drawing.Size(75, 23);
            this.buttonSelectAll.TabIndex = 3;
            this.buttonSelectAll.Text = "Select All";
            this.buttonSelectAll.UseVisualStyleBackColor = true;
            this.buttonSelectAll.Click += new System.EventHandler(this.buttonSelectAll_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.numericUpDownMaxRecord);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.checkBoxRecordLimit);
            this.groupBox2.Location = new System.Drawing.Point(412, 299);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(244, 93);
            this.groupBox2.TabIndex = 21;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Number of Records";
            // 
            // numericUpDownMaxRecord
            // 
            this.numericUpDownMaxRecord.Location = new System.Drawing.Point(88, 25);
            this.numericUpDownMaxRecord.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numericUpDownMaxRecord.Name = "numericUpDownMaxRecord";
            this.numericUpDownMaxRecord.Size = new System.Drawing.Size(121, 20);
            this.numericUpDownMaxRecord.TabIndex = 11;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 27);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Maximum:";
            // 
            // checkBoxRecordLimit
            // 
            this.checkBoxRecordLimit.AutoSize = true;
            this.checkBoxRecordLimit.Location = new System.Drawing.Point(145, 59);
            this.checkBoxRecordLimit.Name = "checkBoxRecordLimit";
            this.checkBoxRecordLimit.Size = new System.Drawing.Size(64, 17);
            this.checkBoxRecordLimit.TabIndex = 12;
            this.checkBoxRecordLimit.Text = "No Limit";
            this.checkBoxRecordLimit.UseVisualStyleBackColor = true;
            this.checkBoxRecordLimit.CheckedChanged += new System.EventHandler(this.checkBoxRecordLimit_CheckedChanged);
            // 
            // FormImportActivities
            // 
            this.AcceptButton = this.buttonImportData;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(680, 457);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonImportData);
            this.Controls.Add(this.buttonImportAndGenerate);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.Name = "FormImportActivities";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Activity Data";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownMaxRecord)).EndInit();
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
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBoxDays;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button buttonUnselectAll;
        private System.Windows.Forms.Button buttonSelectAll;
        private System.Windows.Forms.Button buttonDelete;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.ListBox listBoxSelected;
        private System.Windows.Forms.ListBox listBoxTotal;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.NumericUpDown numericUpDownMaxRecord;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox checkBoxRecordLimit;
        private System.Windows.Forms.CheckBox checkBoxDateLimit;
    }
}