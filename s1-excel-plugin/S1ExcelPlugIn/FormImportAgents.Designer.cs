namespace S1ExcelPlugIn
{
    partial class FormImportAgents
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormImportAgents));
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.buttonImportAndGenerate = new System.Windows.Forms.Button();
            this.buttonImportData = new System.Windows.Forms.Button();
            this.textBoxQuery = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBoxAll = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(307, 199);
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
            this.labelStatus.Location = new System.Drawing.Point(21, 219);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(0, 13);
            this.labelStatus.TabIndex = 6;
            // 
            // buttonImportAndGenerate
            // 
            this.buttonImportAndGenerate.Location = new System.Drawing.Point(105, 199);
            this.buttonImportAndGenerate.Name = "buttonImportAndGenerate";
            this.buttonImportAndGenerate.Size = new System.Drawing.Size(187, 23);
            this.buttonImportAndGenerate.TabIndex = 0;
            this.buttonImportAndGenerate.Tag = "";
            this.buttonImportAndGenerate.Text = "Import Data and Generate Reports";
            this.buttonImportAndGenerate.UseVisualStyleBackColor = true;
            this.buttonImportAndGenerate.Click += new System.EventHandler(this.buttonImportAndGenerate_Click);
            // 
            // buttonImportData
            // 
            this.buttonImportData.Location = new System.Drawing.Point(24, 199);
            this.buttonImportData.Name = "buttonImportData";
            this.buttonImportData.Size = new System.Drawing.Size(75, 23);
            this.buttonImportData.TabIndex = 14;
            this.buttonImportData.Text = "Import Data";
            this.buttonImportData.UseVisualStyleBackColor = true;
            this.buttonImportData.Visible = false;
            this.buttonImportData.Click += new System.EventHandler(this.buttonImportData_Click);
            // 
            // textBoxQuery
            // 
            this.textBoxQuery.Location = new System.Drawing.Point(14, 73);
            this.textBoxQuery.Name = "textBoxQuery";
            this.textBoxQuery.Size = new System.Drawing.Size(417, 20);
            this.textBoxQuery.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 26);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(370, 13);
            this.label2.TabIndex = 17;
            this.label2.Text = "Free text search on Agent\'s network computer name, IP and MAC addresses,";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(358, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Domain and Group name, and agent version (query performed as contains)";
            // 
            // checkBoxAll
            // 
            this.checkBoxAll.AutoSize = true;
            this.checkBoxAll.Location = new System.Drawing.Point(15, 114);
            this.checkBoxAll.Name = "checkBoxAll";
            this.checkBoxAll.Size = new System.Drawing.Size(122, 17);
            this.checkBoxAll.TabIndex = 3;
            this.checkBoxAll.Text = "Return All Endpoints";
            this.checkBoxAll.UseVisualStyleBackColor = true;
            this.checkBoxAll.CheckedChanged += new System.EventHandler(this.checkBoxAll_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.checkBoxAll);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBoxQuery);
            this.groupBox1.Location = new System.Drawing.Point(24, 23);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(447, 155);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Search filter";
            // 
            // FormImportAgents
            // 
            this.AcceptButton = this.buttonImportAndGenerate;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(493, 241);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonImportData);
            this.Controls.Add(this.buttonImportAndGenerate);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.Name = "FormImportAgents";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Agent Data";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button buttonImportAndGenerate;
        private System.Windows.Forms.Button buttonImportData;
        private System.Windows.Forms.TextBox textBoxQuery;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox checkBoxAll;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}