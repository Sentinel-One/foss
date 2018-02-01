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
    public partial class FormAPI : Form
    {
        public FormAPI()
        {
            InitializeComponent();
            syntaxEditURL.Text = Globals.ApiUrl;
            syntaxEditResults.Text = Globals.ApiResults;
        }

        private void FormAPI_Load(object sender, EventArgs e)
        {
            syntaxEditURL.Text = Globals.ApiUrl;
            syntaxEditResults.Text = Globals.ApiResults;
        }
    }
}
