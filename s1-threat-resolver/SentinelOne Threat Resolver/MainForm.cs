using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;


namespace SentinelOne_Threat_Resolver
{
    public partial class MainForm : Form
    {
        string token;

        public MainForm()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }

        private void buttonLogin_Click(object sender, EventArgs e)
        {
            try
            {
                string resourceString = "https://" + textBoxServer.Text.TrimEnd('/') + "/web/api/v1.6/users/login";
                var restClient = new RestClientInterface(
                    endpoint: resourceString,
                    method: HttpVerb.POST,
                    postData: "{\"username\":\"" + textBoxUsername.Text + "\", \"password\":\"" + textBoxPassword.Text + "\"}");
                var results = restClient.MakeRequest();

                dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(results);
                token = x.token;

                toolStripStatusLabel1.Text = "Login successful";
                buttonRetrieve.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

        }

        private void buttonRetrieve_Click(object sender, EventArgs e)
        {
            RetrieveThreats();
        }

        private void RetrieveThreats()
        {
            if (token == "")
            {
                MessageBox.Show("Please login", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            dataGridView1.Rows.Clear();
            toolStripStatusLabel1.Text = "Retrieving data, please wait...";

            dynamic threats;
            int rowCount = 0;

            var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
            {
                NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
            };

            string mitigation_status = comboBox1.SelectedItem.ToString();
            string filter = "";

            switch (mitigation_status)
            {
                case "All":
                    filter = "";
                    break;
                case "Active":
                    filter = "&mitigation_status=1";
                    break;
                case "Mitigated":
                    filter = "&mitigation_status=0";
                    break;
                case "Blocked":
                    filter = "&mitigation_status=2";
                    break;
                case "Suspicious":
                    filter = "&mitigation_status=3";
                    break;
            }

            string resourceString = "https://" + textBoxServer.Text + "/web/api/v1.6/threats?resolved=false&limit=10000" + filter;
            var restClient = new RestClientInterface(resourceString);
            restClient.EndPoint = resourceString;
            restClient.Method = HttpVerb.GET;
            var results = restClient.MakeRequest(token).ToString();
            threats = Newtonsoft.Json.JsonConvert.DeserializeObject(results, JsonSettings);
            rowCount = (int)threats.Count;
            string miti_status = "";
            string status = "";

            if (rowCount > 0)
            {
                for (int i = 0; i < rowCount; i++)
                {
                    status = threats[i].mitigation_status;

                    switch (status)
                    {
                        case "0":
                            miti_status = "Mitigated";
                            break;
                        case "1":
                            miti_status = "Active";
                            break;
                        case "2":
                            miti_status = "Blocked";
                            break;
                        case "3":
                            miti_status = "Suspicious";
                            break;
                        case "4":
                            miti_status = "Pending";
                            break;
                        case "5":
                            miti_status = "Suspicious canceled";
                            break;
                    }

                    dataGridView1.Rows.Add(null,
                            threats[i].id,
                            miti_status,
                            threats[i].agent,
                            threats[i].created_date,
                            threats[i].file_id.display_name);
                }

                buttonResolve.Enabled = true;
                buttonSelectAll.Enabled = true;
                buttonUnselect.Enabled = true;

                if (rowCount == 1)
                {
                    toolStripStatusLabel1.Text = "There is " + rowCount.ToString() + " unresolved threat";
                }
                else
                {
                    toolStripStatusLabel1.Text = "There are " + rowCount.ToString() + " unresolved threats";
                }
            }
            else
            {
                toolStripStatusLabel1.Text = "No unresolved threats found";
            }

        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                chk.Value = true;
            }
            toolStripStatusLabel1.Text = dataGridView1.Rows.Count.ToString() + " rows selected";
        }

        private void buttonUnselect_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[0];
                chk.Value = false;
            }
            toolStripStatusLabel1.Text = "Unselected all rows";

        }

        private void buttonResolve_Click(object sender, EventArgs e)
        {
            int rowsProcessed = 0;
            int rowCount = dataGridView1.Rows.Count;
            string resourceString = "";
            var restClient = new RestClientInterface();
            var results = "";

            toolStripStatusLabel1.Text = "Resolving threats, please wait...";
            statusStrip1.Update();

            for (int i = dataGridView1.Rows.Count; i --> 0;)
            {
                if (Convert.ToBoolean( dataGridView1.Rows[i].Cells[0].Value ) == true)
                {
                    resourceString = "https://" + textBoxServer.Text + "/web/api/v1.6/threats/" + dataGridView1.Rows[i].Cells["ThreatID"].Value + "/resolve";
                    // var restClient = new RestClientInterface(resourceString);
                    restClient.EndPoint = resourceString;
                    restClient.Method = HttpVerb.POST;
                    results = restClient.MakeRequest(token).ToString();
                    // dataGridView1.Rows.RemoveAt(i);
                    rowsProcessed++;
                }
            }
            toolStripStatusLabel1.Text = "Resolved " + rowsProcessed.ToString() + " threats, refreshing data from server...";
            statusStrip1.Update();
            RetrieveThreats();
        }
    }
}
