namespace S1ExcelPlugIn
{
    partial class FormAPI
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormAPI));
            this.syntaxEditURL = new QWhale.Editor.SyntaxEdit(this.components);
            this.syntaxEditResults = new QWhale.Editor.SyntaxEdit(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.parserURL = new QWhale.Syntax.Parser();
            this.parserJSON = new QWhale.Syntax.Parser();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // syntaxEditURL
            // 
            this.syntaxEditURL.BackColor = System.Drawing.SystemColors.Window;
            this.syntaxEditURL.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.syntaxEditURL.Dock = System.Windows.Forms.DockStyle.Fill;
            this.syntaxEditURL.Font = new System.Drawing.Font("Courier New", 10F);
            this.syntaxEditURL.Lexer = this.parserURL;
            this.syntaxEditURL.Location = new System.Drawing.Point(3, 16);
            this.syntaxEditURL.Name = "syntaxEditURL";
            this.syntaxEditURL.Size = new System.Drawing.Size(674, 81);
            this.syntaxEditURL.TabIndex = 0;
            this.syntaxEditURL.Text = "";
            this.syntaxEditURL.WordWrap = true;
            // 
            // syntaxEditResults
            // 
            this.syntaxEditResults.BackColor = System.Drawing.SystemColors.Window;
            this.syntaxEditResults.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.syntaxEditResults.Dock = System.Windows.Forms.DockStyle.Fill;
            this.syntaxEditResults.Font = new System.Drawing.Font("Courier New", 10F);
            this.syntaxEditResults.Gutter.Options = ((QWhale.Editor.GutterOptions)(((QWhale.Editor.GutterOptions.PaintLineNumbers | QWhale.Editor.GutterOptions.PaintBookMarks) 
            | QWhale.Editor.GutterOptions.PaintLineModificators)));
            this.syntaxEditResults.Lexer = this.parserJSON;
            this.syntaxEditResults.Location = new System.Drawing.Point(3, 16);
            this.syntaxEditResults.Name = "syntaxEditResults";
            this.syntaxEditResults.Size = new System.Drawing.Size(671, 292);
            this.syntaxEditResults.TabIndex = 1;
            this.syntaxEditResults.Text = "";
            this.syntaxEditResults.WordWrap = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.syntaxEditURL);
            this.groupBox1.Location = new System.Drawing.Point(6, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(680, 100);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Last REST API Call";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.syntaxEditResults);
            this.groupBox2.Location = new System.Drawing.Point(9, 118);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(677, 311);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Last REST API Results";
            // 
            // parserURL
            // 
            this.parserURL.DefaultState = 0;
            this.parserURL.XmlScheme = resources.GetString("parserURL.XmlScheme");
            // 
            // parserJSON
            // 
            this.parserJSON.DefaultState = 0;
            this.parserJSON.XmlScheme = resources.GetString("parserJSON.XmlScheme");
            // 
            // FormAPI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(695, 441);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormAPI";
            this.Text = "Show API Details";
            this.Load += new System.EventHandler(this.FormAPI_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private QWhale.Editor.SyntaxEdit syntaxEditURL;
        private QWhale.Editor.SyntaxEdit syntaxEditResults;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private QWhale.Syntax.Parser parserURL;
        private QWhale.Syntax.Parser parserJSON;
    }
}