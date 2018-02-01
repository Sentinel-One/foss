namespace S1ExcelPlugIn
{
    partial class FormLogin
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormLogin));
            this.a1Panel2 = new Owf.Controls.A1Panel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.checkBoxToken = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxToken = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBoxUsername = new System.Windows.Forms.ComboBox();
            this.comboBoxURL = new System.Windows.Forms.ComboBox();
            this.buttonClearAll = new System.Windows.Forms.Button();
            this.buttonClearOne = new System.Windows.Forms.Button();
            this.labelLoginMessage = new System.Windows.Forms.Label();
            this.buttonLogin = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.a1Panel2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // a1Panel2
            // 
            this.a1Panel2.BackColor = System.Drawing.Color.Transparent;
            this.a1Panel2.BorderColor = System.Drawing.Color.Gray;
            this.a1Panel2.BorderWidth = 0;
            this.a1Panel2.Controls.Add(this.groupBox1);
            this.a1Panel2.Controls.Add(this.buttonClearAll);
            this.a1Panel2.Controls.Add(this.buttonClearOne);
            this.a1Panel2.Controls.Add(this.labelLoginMessage);
            this.a1Panel2.Controls.Add(this.buttonLogin);
            this.a1Panel2.Controls.Add(this.buttonCancel);
            this.a1Panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.a1Panel2.GradientEndColor = System.Drawing.Color.Silver;
            this.a1Panel2.GradientStartColor = System.Drawing.Color.White;
            this.a1Panel2.Image = null;
            this.a1Panel2.ImageLocation = new System.Drawing.Point(4, 4);
            this.a1Panel2.Location = new System.Drawing.Point(0, 0);
            this.a1Panel2.Name = "a1Panel2";
            this.a1Panel2.ShadowOffSet = 0;
            this.a1Panel2.Size = new System.Drawing.Size(434, 344);
            this.a1Panel2.TabIndex = 9;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBoxPassword);
            this.groupBox1.Controls.Add(this.checkBoxToken);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBoxToken);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.comboBoxUsername);
            this.groupBox1.Controls.Add(this.comboBoxURL);
            this.groupBox1.Location = new System.Drawing.Point(12, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(406, 246);
            this.groupBox1.TabIndex = 19;
            this.groupBox1.TabStop = false;
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxPassword.Location = new System.Drawing.Point(139, 105);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.PasswordChar = '*';
            this.textBoxPassword.Size = new System.Drawing.Size(248, 22);
            this.textBoxPassword.TabIndex = 3;
            // 
            // checkBoxToken
            // 
            this.checkBoxToken.AutoSize = true;
            this.checkBoxToken.Location = new System.Drawing.Point(308, 140);
            this.checkBoxToken.Name = "checkBoxToken";
            this.checkBoxToken.Size = new System.Drawing.Size(79, 17);
            this.checkBoxToken.TabIndex = 18;
            this.checkBoxToken.Text = "Use Token";
            this.checkBoxToken.UseVisualStyleBackColor = true;
            this.checkBoxToken.CheckedChanged += new System.EventHandler(this.checkBoxToken_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(10, 108);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 14);
            this.label3.TabIndex = 4;
            this.label3.Text = "Password:";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(11, 166);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(46, 14);
            this.label6.TabIndex = 17;
            this.label6.Text = "Token:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(10, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 14);
            this.label1.TabIndex = 1;
            this.label1.Text = "Management Server";
            // 
            // textBoxToken
            // 
            this.textBoxToken.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxToken.Location = new System.Drawing.Point(139, 163);
            this.textBoxToken.Multiline = true;
            this.textBoxToken.Name = "textBoxToken";
            this.textBoxToken.Size = new System.Drawing.Size(248, 61);
            this.textBoxToken.TabIndex = 16;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(10, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 14);
            this.label2.TabIndex = 3;
            this.label2.Text = "Username:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(137, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(223, 14);
            this.label5.TabIndex = 11;
            this.label5.Text = "Example: https://name.sentinelone.net";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(12, 44);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(32, 14);
            this.label4.TabIndex = 12;
            this.label4.Text = "URL:";
            // 
            // comboBoxUsername
            // 
            this.comboBoxUsername.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxUsername.FormattingEnabled = true;
            this.comboBoxUsername.Location = new System.Drawing.Point(139, 73);
            this.comboBoxUsername.Name = "comboBoxUsername";
            this.comboBoxUsername.Size = new System.Drawing.Size(248, 24);
            this.comboBoxUsername.TabIndex = 2;
            this.comboBoxUsername.SelectedValueChanged += new System.EventHandler(this.comboBoxUsername_SelectedValueChanged);
            // 
            // comboBoxURL
            // 
            this.comboBoxURL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxURL.FormattingEnabled = true;
            this.comboBoxURL.Location = new System.Drawing.Point(139, 22);
            this.comboBoxURL.Name = "comboBoxURL";
            this.comboBoxURL.Size = new System.Drawing.Size(248, 24);
            this.comboBoxURL.Sorted = true;
            this.comboBoxURL.TabIndex = 1;
            this.comboBoxURL.SelectedValueChanged += new System.EventHandler(this.comboBoxURL_SelectedValueChanged);
            // 
            // buttonClearAll
            // 
            this.buttonClearAll.Location = new System.Drawing.Point(289, 305);
            this.buttonClearAll.Name = "buttonClearAll";
            this.buttonClearAll.Size = new System.Drawing.Size(111, 23);
            this.buttonClearAll.TabIndex = 15;
            this.buttonClearAll.Text = "Clear All Credentials";
            this.buttonClearAll.UseVisualStyleBackColor = true;
            this.buttonClearAll.Click += new System.EventHandler(this.buttonClearAll_Click);
            // 
            // buttonClearOne
            // 
            this.buttonClearOne.Location = new System.Drawing.Point(152, 305);
            this.buttonClearOne.Name = "buttonClearOne";
            this.buttonClearOne.Size = new System.Drawing.Size(116, 23);
            this.buttonClearOne.TabIndex = 14;
            this.buttonClearOne.Text = "Clear Credential";
            this.buttonClearOne.UseVisualStyleBackColor = true;
            this.buttonClearOne.Click += new System.EventHandler(this.buttonClearOne_Click);
            // 
            // labelLoginMessage
            // 
            this.labelLoginMessage.AutoSize = true;
            this.labelLoginMessage.BackColor = System.Drawing.Color.Transparent;
            this.labelLoginMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelLoginMessage.ForeColor = System.Drawing.Color.ForestGreen;
            this.labelLoginMessage.Location = new System.Drawing.Point(24, 270);
            this.labelLoginMessage.Name = "labelLoginMessage";
            this.labelLoginMessage.Size = new System.Drawing.Size(116, 15);
            this.labelLoginMessage.TabIndex = 13;
            this.labelLoginMessage.Text = "Login Successful";
            this.labelLoginMessage.Visible = false;
            // 
            // buttonLogin
            // 
            this.buttonLogin.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonLogin.Location = new System.Drawing.Point(152, 267);
            this.buttonLogin.Name = "buttonLogin";
            this.buttonLogin.Size = new System.Drawing.Size(116, 23);
            this.buttonLogin.TabIndex = 4;
            this.buttonLogin.Text = "Login";
            this.buttonLogin.UseVisualStyleBackColor = true;
            this.buttonLogin.Click += new System.EventHandler(this.buttonLogin_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(289, 267);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(111, 23);
            this.buttonCancel.TabIndex = 5;
            this.buttonCancel.Text = "Close";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // FormLogin
            // 
            this.AcceptButton = this.buttonLogin;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 344);
            this.Controls.Add(this.a1Panel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormLogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Login";
            this.Load += new System.EventHandler(this.FormConfig_Load);
            this.a1Panel2.ResumeLayout(false);
            this.a1Panel2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxPassword;
        private Owf.Controls.A1Panel a1Panel2;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Button buttonLogin;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label labelLoginMessage;
        private System.Windows.Forms.ComboBox comboBoxUsername;
        private System.Windows.Forms.ComboBox comboBoxURL;
        private System.Windows.Forms.Button buttonClearAll;
        private System.Windows.Forms.Button buttonClearOne;
        private System.Windows.Forms.CheckBox checkBoxToken;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBoxToken;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}