namespace S1ExcelPlugIn
{
    partial class FormImportProcesses
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
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.buttonImportAndGenerate = new System.Windows.Forms.Button();
            this.buttonImportData = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(369, 54);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(54, 23);
            this.buttonClose.TabIndex = 3;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // labelStatus
            // 
            this.labelStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelStatus.AutoSize = true;
            this.labelStatus.Location = new System.Drawing.Point(21, 98);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(0, 13);
            this.labelStatus.TabIndex = 6;
            // 
            // buttonImportAndGenerate
            // 
            this.buttonImportAndGenerate.Location = new System.Drawing.Point(160, 54);
            this.buttonImportAndGenerate.Name = "buttonImportAndGenerate";
            this.buttonImportAndGenerate.Size = new System.Drawing.Size(187, 23);
            this.buttonImportAndGenerate.TabIndex = 8;
            this.buttonImportAndGenerate.Tag = "";
            this.buttonImportAndGenerate.Text = "Import Data and Generate Reports";
            this.buttonImportAndGenerate.UseVisualStyleBackColor = true;
            this.buttonImportAndGenerate.Click += new System.EventHandler(this.buttonImportAndGenerate_Click);
            // 
            // buttonImportData
            // 
            this.buttonImportData.Location = new System.Drawing.Point(64, 54);
            this.buttonImportData.Name = "buttonImportData";
            this.buttonImportData.Size = new System.Drawing.Size(75, 23);
            this.buttonImportData.TabIndex = 14;
            this.buttonImportData.Text = "Import Data";
            this.buttonImportData.UseVisualStyleBackColor = true;
            this.buttonImportData.Click += new System.EventHandler(this.buttonImportData_Click);
            // 
            // FormImportProcesses
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 120);
            this.Controls.Add(this.buttonImportData);
            this.Controls.Add(this.buttonImportAndGenerate);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MinimizeBox = false;
            this.Name = "FormImportProcesses";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Process Data";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button buttonImportAndGenerate;
        private System.Windows.Forms.Button buttonImportData;
    }
}