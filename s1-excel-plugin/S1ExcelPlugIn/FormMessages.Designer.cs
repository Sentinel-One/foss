namespace S1ExcelPlugIn
{
    partial class FormMessages
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
            this.labelMessage1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.labelMessage2 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelMessage1
            // 
            this.labelMessage1.AutoSize = true;
            this.labelMessage1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMessage1.Location = new System.Drawing.Point(22, 25);
            this.labelMessage1.Name = "labelMessage1";
            this.labelMessage1.Size = new System.Drawing.Size(108, 18);
            this.labelMessage1.TabIndex = 0;
            this.labelMessage1.Text = "Processing...";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.buttonCancel);
            this.groupBox1.Controls.Add(this.labelMessage2);
            this.groupBox1.Controls.Add(this.labelMessage1);
            this.groupBox1.Location = new System.Drawing.Point(6, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(608, 83);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(522, 49);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 2;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // labelMessage2
            // 
            this.labelMessage2.AutoSize = true;
            this.labelMessage2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelMessage2.Location = new System.Drawing.Point(23, 47);
            this.labelMessage2.Name = "labelMessage2";
            this.labelMessage2.Size = new System.Drawing.Size(0, 18);
            this.labelMessage2.TabIndex = 1;
            // 
            // FormMessages
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(620, 90);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FormMessages";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormMessages";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label labelMessage1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label labelMessage2;
        private System.Windows.Forms.Button buttonCancel;
    }
}