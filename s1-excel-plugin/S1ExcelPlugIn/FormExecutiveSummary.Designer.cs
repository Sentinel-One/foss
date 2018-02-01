namespace S1ExcelPlugIn
{
    partial class FormExecutiveSummary
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormExecutiveSummary));
            this.buttonReport = new System.Windows.Forms.Button();
            this.checkBoxHTML = new System.Windows.Forms.CheckBox();
            this.checkBoxPDF = new System.Windows.Forms.CheckBox();
            this.checkBoxEmail = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.buttonClose = new System.Windows.Forms.Button();
            this.labelEmailMessage = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.textBoxBody = new System.Windows.Forms.TextBox();
            this.textBoxMailTo = new System.Windows.Forms.TextBox();
            this.textBoxSubject = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label1TestMessage = new System.Windows.Forms.Label();
            this.buttonTest = new System.Windows.Forms.Button();
            this.checkBoxSSL = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxMailFrom = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxUsername = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxSMTPPort = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBoxSMTPHost = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonReport
            // 
            this.buttonReport.Location = new System.Drawing.Point(209, 74);
            this.buttonReport.Name = "buttonReport";
            this.buttonReport.Size = new System.Drawing.Size(90, 23);
            this.buttonReport.TabIndex = 1;
            this.buttonReport.Text = "Generate";
            this.buttonReport.UseVisualStyleBackColor = true;
            this.buttonReport.Click += new System.EventHandler(this.buttonReport_Click);
            // 
            // checkBoxHTML
            // 
            this.checkBoxHTML.AutoSize = true;
            this.checkBoxHTML.Checked = true;
            this.checkBoxHTML.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxHTML.Location = new System.Drawing.Point(238, 19);
            this.checkBoxHTML.Name = "checkBoxHTML";
            this.checkBoxHTML.Size = new System.Drawing.Size(56, 17);
            this.checkBoxHTML.TabIndex = 2;
            this.checkBoxHTML.Text = "HTML";
            this.checkBoxHTML.UseVisualStyleBackColor = true;
            // 
            // checkBoxPDF
            // 
            this.checkBoxPDF.AutoSize = true;
            this.checkBoxPDF.Checked = true;
            this.checkBoxPDF.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxPDF.Location = new System.Drawing.Point(322, 19);
            this.checkBoxPDF.Name = "checkBoxPDF";
            this.checkBoxPDF.Size = new System.Drawing.Size(47, 17);
            this.checkBoxPDF.TabIndex = 3;
            this.checkBoxPDF.Text = "PDF";
            this.checkBoxPDF.UseVisualStyleBackColor = true;
            // 
            // checkBoxEmail
            // 
            this.checkBoxEmail.AutoSize = true;
            this.checkBoxEmail.Checked = true;
            this.checkBoxEmail.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxEmail.Location = new System.Drawing.Point(27, 90);
            this.checkBoxEmail.Name = "checkBoxEmail";
            this.checkBoxEmail.Size = new System.Drawing.Size(133, 17);
            this.checkBoxEmail.TabIndex = 4;
            this.checkBoxEmail.Text = "Email as Attachment(s)";
            this.checkBoxEmail.UseVisualStyleBackColor = true;
            this.checkBoxEmail.CheckedChanged += new System.EventHandler(this.checkBoxEmail_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBoxHTML);
            this.groupBox1.Controls.Add(this.checkBoxPDF);
            this.groupBox1.Location = new System.Drawing.Point(24, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(389, 47);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Output Formats";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(442, 351);
            this.tabControl1.TabIndex = 8;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.buttonClose);
            this.tabPage1.Controls.Add(this.labelEmailMessage);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.buttonReport);
            this.tabPage1.Controls.Add(this.checkBoxEmail);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(434, 325);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Generate Report";
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(319, 74);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(74, 23);
            this.buttonClose.TabIndex = 15;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // labelEmailMessage
            // 
            this.labelEmailMessage.AutoSize = true;
            this.labelEmailMessage.BackColor = System.Drawing.Color.Transparent;
            this.labelEmailMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelEmailMessage.ForeColor = System.Drawing.Color.ForestGreen;
            this.labelEmailMessage.Location = new System.Drawing.Point(24, 296);
            this.labelEmailMessage.Name = "labelEmailMessage";
            this.labelEmailMessage.Size = new System.Drawing.Size(156, 15);
            this.labelEmailMessage.TabIndex = 14;
            this.labelEmailMessage.Text = "Email sent successfully";
            this.labelEmailMessage.Visible = false;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(24, 277);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(0, 13);
            this.label10.TabIndex = 13;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBoxBody);
            this.groupBox2.Controls.Add(this.textBoxMailTo);
            this.groupBox2.Controls.Add(this.textBoxSubject);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Location = new System.Drawing.Point(24, 103);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(389, 187);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            // 
            // textBoxBody
            // 
            this.textBoxBody.Location = new System.Drawing.Point(101, 122);
            this.textBoxBody.Multiline = true;
            this.textBoxBody.Name = "textBoxBody";
            this.textBoxBody.Size = new System.Drawing.Size(266, 52);
            this.textBoxBody.TabIndex = 11;
            this.textBoxBody.Text = "Please see attached report";
            // 
            // textBoxMailTo
            // 
            this.textBoxMailTo.Location = new System.Drawing.Point(101, 22);
            this.textBoxMailTo.Multiline = true;
            this.textBoxMailTo.Name = "textBoxMailTo";
            this.textBoxMailTo.Size = new System.Drawing.Size(268, 36);
            this.textBoxMailTo.TabIndex = 9;
            // 
            // textBoxSubject
            // 
            this.textBoxSubject.Location = new System.Drawing.Point(101, 68);
            this.textBoxSubject.Multiline = true;
            this.textBoxSubject.Name = "textBoxSubject";
            this.textBoxSubject.Size = new System.Drawing.Size(268, 45);
            this.textBoxSubject.TabIndex = 10;
            this.textBoxSubject.Text = "SentinelOne Server Summary Report Attached";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(17, 25);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(45, 13);
            this.label7.TabIndex = 6;
            this.label7.Text = "Mail To:";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(17, 71);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(46, 13);
            this.label8.TabIndex = 7;
            this.label8.Text = "Subject:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(17, 122);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(80, 13);
            this.label9.TabIndex = 8;
            this.label9.Text = "Message Body:";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(434, 325);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Configure Email";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label1TestMessage);
            this.groupBox3.Controls.Add(this.buttonTest);
            this.groupBox3.Controls.Add(this.checkBoxSSL);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.textBoxMailFrom);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.textBoxPassword);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.textBoxUsername);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.textBoxSMTPPort);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.textBoxSMTPHost);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Location = new System.Drawing.Point(19, 15);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(395, 274);
            this.groupBox3.TabIndex = 9;
            this.groupBox3.TabStop = false;
            // 
            // label1TestMessage
            // 
            this.label1TestMessage.AutoSize = true;
            this.label1TestMessage.BackColor = System.Drawing.Color.Transparent;
            this.label1TestMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1TestMessage.ForeColor = System.Drawing.Color.ForestGreen;
            this.label1TestMessage.Location = new System.Drawing.Point(18, 239);
            this.label1TestMessage.Name = "label1TestMessage";
            this.label1TestMessage.Size = new System.Drawing.Size(186, 15);
            this.label1TestMessage.TabIndex = 15;
            this.label1TestMessage.Text = "Test email sent successfully";
            this.label1TestMessage.Visible = false;
            // 
            // buttonTest
            // 
            this.buttonTest.Location = new System.Drawing.Point(305, 236);
            this.buttonTest.Name = "buttonTest";
            this.buttonTest.Size = new System.Drawing.Size(75, 23);
            this.buttonTest.TabIndex = 13;
            this.buttonTest.Text = "Test";
            this.buttonTest.UseVisualStyleBackColor = true;
            this.buttonTest.Click += new System.EventHandler(this.buttonTest_Click);
            // 
            // checkBoxSSL
            // 
            this.checkBoxSSL.AutoSize = true;
            this.checkBoxSSL.Checked = true;
            this.checkBoxSSL.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxSSL.Location = new System.Drawing.Point(129, 100);
            this.checkBoxSSL.Name = "checkBoxSSL";
            this.checkBoxSSL.Size = new System.Drawing.Size(15, 14);
            this.checkBoxSSL.TabIndex = 12;
            this.checkBoxSSL.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "SMTP Host:";
            // 
            // textBoxMailFrom
            // 
            this.textBoxMailFrom.Location = new System.Drawing.Point(129, 202);
            this.textBoxMailFrom.Name = "textBoxMailFrom";
            this.textBoxMailFrom.Size = new System.Drawing.Size(251, 20);
            this.textBoxMailFrom.TabIndex = 11;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(18, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "SMTP Port:";
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(129, 166);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.PasswordChar = '*';
            this.textBoxPassword.Size = new System.Drawing.Size(251, 20);
            this.textBoxPassword.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(18, 100);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(99, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "SMTP Enable SSL:";
            // 
            // textBoxUsername
            // 
            this.textBoxUsername.Location = new System.Drawing.Point(129, 130);
            this.textBoxUsername.Name = "textBoxUsername";
            this.textBoxUsername.Size = new System.Drawing.Size(251, 20);
            this.textBoxUsername.TabIndex = 9;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(18, 133);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "Username:";
            // 
            // textBoxSMTPPort
            // 
            this.textBoxSMTPPort.Location = new System.Drawing.Point(129, 64);
            this.textBoxSMTPPort.Name = "textBoxSMTPPort";
            this.textBoxSMTPPort.Size = new System.Drawing.Size(251, 20);
            this.textBoxSMTPPort.TabIndex = 7;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(18, 169);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(56, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Password:";
            // 
            // textBoxSMTPHost
            // 
            this.textBoxSMTPHost.Location = new System.Drawing.Point(129, 28);
            this.textBoxSMTPHost.Name = "textBoxSMTPHost";
            this.textBoxSMTPHost.Size = new System.Drawing.Size(251, 20);
            this.textBoxSMTPHost.TabIndex = 6;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(18, 205);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(55, 13);
            this.label6.TabIndex = 5;
            this.label6.Text = "Mail From:";
            // 
            // FormExecutiveSummary
            // 
            this.AcceptButton = this.buttonReport;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 375);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormExecutiveSummary";
            this.ShowIcon = false;
            this.Text = "Executive Summary";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormExecutiveSummary_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button buttonReport;
        private System.Windows.Forms.CheckBox checkBoxHTML;
        private System.Windows.Forms.CheckBox checkBoxPDF;
        private System.Windows.Forms.CheckBox checkBoxEmail;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox textBoxBody;
        private System.Windows.Forms.TextBox textBoxMailTo;
        private System.Windows.Forms.TextBox textBoxSubject;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textBoxMailFrom;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.TextBox textBoxUsername;
        private System.Windows.Forms.TextBox textBoxSMTPPort;
        private System.Windows.Forms.TextBox textBoxSMTPHost;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBoxSSL;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label labelEmailMessage;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Button buttonTest;
        private System.Windows.Forms.Label label1TestMessage;
    }
}