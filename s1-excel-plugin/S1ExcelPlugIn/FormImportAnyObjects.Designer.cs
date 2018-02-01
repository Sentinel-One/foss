namespace S1ExcelPlugIn
{
    partial class FormImportAnyObjects
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormImportAnyObjects));
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelStatus = new System.Windows.Forms.Label();
            this.buttonImportData = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.listBoxObjects = new System.Windows.Forms.ListBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(176, 268);
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
            this.labelStatus.Location = new System.Drawing.Point(21, 286);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(0, 13);
            this.labelStatus.TabIndex = 6;
            // 
            // buttonImportData
            // 
            this.buttonImportData.Location = new System.Drawing.Point(78, 268);
            this.buttonImportData.Name = "buttonImportData";
            this.buttonImportData.Size = new System.Drawing.Size(75, 23);
            this.buttonImportData.TabIndex = 0;
            this.buttonImportData.Text = "Import Data";
            this.buttonImportData.UseVisualStyleBackColor = true;
            this.buttonImportData.Click += new System.EventHandler(this.buttonImportData_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listBoxObjects);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(295, 250);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Select an object to import";
            // 
            // listBoxObjects
            // 
            this.listBoxObjects.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxObjects.FormattingEnabled = true;
            this.listBoxObjects.Items.AddRange(new object[] {
            "Activities/Types",
            "Agents/Configuration",
            "Agents/Passphrase",
            "Exclusion-Lists",
            "Filters",
            "Hashes",
            "Reports",
            "Users"});
            this.listBoxObjects.Location = new System.Drawing.Point(3, 16);
            this.listBoxObjects.Name = "listBoxObjects";
            this.listBoxObjects.Size = new System.Drawing.Size(289, 231);
            this.listBoxObjects.TabIndex = 2;
            // 
            // FormImportAnyObjects
            // 
            this.AcceptButton = this.buttonImportData;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(318, 308);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonImportData);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.buttonClose);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.Name = "FormImportAnyObjects";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Object";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button buttonImportData;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListBox listBoxObjects;
    }
}