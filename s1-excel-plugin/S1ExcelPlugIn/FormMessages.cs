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
    public partial class FormMessages : Form
    {
        public bool StopProcessing = false;

        public FormMessages()
        {
            InitializeComponent();
            labelMessage1.Update();
        }

        public void UpdateMessage(string msg1, string msg2)
        {
            labelMessage1.Text = msg1;
            labelMessage2.Text = msg2;
            Application.DoEvents();
        }

        public void StartMessage()
        {
            labelMessage1.Text = "Processing...";
            labelMessage2.Text = "";
            Show();
            Application.DoEvents();
        }

        public void Message(string msg1, string msg2, bool allowCancel)
        {
            labelMessage1.Text = msg1;
            labelMessage2.Text = msg2;
            buttonCancel.Visible = allowCancel;
            Show();
            Application.DoEvents();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            StopProcessing = true;
        }
    }
}
