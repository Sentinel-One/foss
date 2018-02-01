using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace S1ExcelPlugIn
{
    public partial class FormPrompt : Form
    {
        public string AuthCode
        {
            get; set;
        }

        public FormPrompt()
        {
            InitializeComponent();

            textBoxVerify.GotFocus += new System.EventHandler(this.textBoxVerify_GotFocus);
            textBoxVerify.LostFocus += new System.EventHandler(this.textBoxVerify_LostFocus);
        }

        private void textBoxVerify_GotFocus(object sender, EventArgs e)
        {
            textBoxVerify.Text = "";
            textBoxVerify.ForeColor = System.Drawing.SystemColors.WindowText;
        }

        private void textBoxVerify_LostFocus(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(textBoxVerify.Text))
            {
                textBoxVerify.Text = "Authentication code";
                textBoxVerify.ForeColor = Color.Gray;
            }
        }

        private void buttonVerify_Click(object sender, EventArgs e)
        {
            this.AuthCode = textBoxVerify.Text;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
