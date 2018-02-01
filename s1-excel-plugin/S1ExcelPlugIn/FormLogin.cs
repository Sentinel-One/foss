using System;
using System.Windows.Forms;
using Microsoft.Win32;
using System.IO;
using System.Security.Cryptography;

namespace S1ExcelPlugIn
{
    public partial class FormLogin : Form
    {
        string userName = "";
        string serverName = "";
        string password = "";
        string token;
        Crypto crypto = new Crypto();
        
        public FormLogin()
        {
            InitializeComponent();
        }

        private void FormConfig_Load(object sender, EventArgs e)
        {
            try
            {
                RegistryKey Credentials = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\ExcelPlugIn\\Credentials");

                string serverName = "";

                checkBoxToken.Checked = false;
                textBoxToken.Enabled = false;

                if (Credentials.GetValueNames().Length > 0)
                {
                    foreach (string valueName in Credentials.GetValueNames())
                    {
                        serverName = valueName.Substring(0, valueName.IndexOf("||"));
                        if (comboBoxURL.Items.Contains(serverName))
                        {
                            // Already in the combo box, no need to add
                        }
                        else
                        {
                            comboBoxURL.Items.Add(serverName);
                        }
                    }

                    comboBoxURL.SelectedText = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                    comboBoxUsername.SelectedText = crypto.GetSettings("ExcelPlugIn", "Username");

                    if (crypto.GetSettings("ExcelPlugIn", "UseToken") == null ||
                        Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "UseToken")) == false)
                    {
                        checkBoxToken.Checked = false;
                        textBoxPassword.Text = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Password"));
                        comboBoxUsername.Enabled = true;
                        textBoxPassword.Enabled = true;
                        textBoxToken.Enabled = false;
                    }
                    else
                    {
                        checkBoxToken.Checked = true;
                        textBoxToken.Text = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                        comboBoxUsername.Enabled = false;
                        textBoxPassword.Enabled = false;
                        textBoxToken.Enabled = true;
                    }


                    if (comboBoxURL.Text != "")
                        this.comboBoxURL_SelectedValueChanged(this, new EventArgs());
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error retrieving settings", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonLogin_Click(object sender, EventArgs e)
        {
            try
            {
                if (checkBoxToken.Checked)
                {
                    string resourceString = comboBoxURL.Text.TrimEnd('/') + "/web/api/v1.6/threats/summary";
                    var restClient = new RestClientInterface(
                        endpoint: resourceString,
                        method: HttpVerb.GET);
                    var results = restClient.MakeRequest(textBoxToken.Text);
                    token = textBoxToken.Text;
                }
                else
                {
                    string resourceString = comboBoxURL.Text.TrimEnd('/') + "/web/api/v1.6/users/login";
                    var restClient = new RestClientInterface(
                        endpoint: resourceString,
                        method: HttpVerb.POST,
                        postData: "{\"username\":\"" + comboBoxUsername.Text + "\", \"password\":\"" + textBoxPassword.Text + "\"}");
                    var results = restClient.MakeRequest();

                    dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(results);
                    token = x.token;
                }

                labelLoginMessage.Visible = true;
                SaveAll();
            }
            catch (Exception ex)
            {
                #region 2FA
                if (ex.Message.Contains("Required2FAConfiguration"))
                {
                    MessageBox.Show("Two factor authentication has been enabled for this account.\r\n\r\n" + "You need to use the SentinelOne management console to configure a two factor authenticator before logging in here.", "2FA Configuration Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
                else if (ex.Message.Contains("Required2FAAutentication"))
                {
                    try
                    {
                        string authCode = "";
                        FormPrompt promptWin = new FormPrompt();
                        promptWin.StartPosition = FormStartPosition.CenterParent;
                        var dialogResult = promptWin.ShowDialog();

                        if (dialogResult == DialogResult.OK)
                        {
                            authCode = promptWin.AuthCode;
                            string authToken = "";
                            dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(ex.Message.Substring(ex.Message.IndexOf("{")));
                            authToken = x.token;
                            string resourceString = comboBoxURL.Text.TrimEnd('/') + "/web/api/v1.6/users/auth/app?token=" + authToken;
                            var restClient = new RestClientInterface(
                                endpoint: resourceString,
                                method: HttpVerb.POST,
                                postData: "{\"code\":\"" + authCode + "\"}");
                            var results = restClient.MakeRequest(authToken);

                            string newToken = "";
                            dynamic y = Newtonsoft.Json.JsonConvert.DeserializeObject(results);
                            newToken = y.token;

                            token = newToken;
                            labelLoginMessage.Visible = true;
                            SaveAll();
                        }
                        else
                        {
                            return;
                        }

                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show(ex2.Message, "2FA Login Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    #endregion
                }
                else
                {
                    MessageBox.Show(ex.Message, "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        #region Saving and Retrieving from Registry
 
        private void SaveAll()
        {
            try
            {
                if (checkBoxToken.Checked)
                {
                    string credName = "";
                    string credValue = "";
                    credName = comboBoxURL.Text + "||" + "Token";
                    credValue = "not_relevant" + "||" + token;

                    crypto.SetSettings("ExcelPlugIn", "ManagementServer", comboBoxURL.Text);
                    crypto.SetSettings("ExcelPlugIn", "Username", "Token");
                    crypto.SetSettings("ExcelPlugIn", "Password", crypto.Encrypt("not_relevant"));
                    crypto.SetSettings("ExcelPlugIn", "Token", crypto.Encrypt(token));
                    crypto.SetSettings("ExcelPlugIn", "UseToken", checkBoxToken.Checked.ToString());

                    crypto.SetSettings("ExcelPlugIn\\Credentials", credName, crypto.Encrypt(credValue));
                }
                else
                {
                    string credName = "";
                    string credValue = "";
                    credName = comboBoxURL.Text + "||" + comboBoxUsername.Text;
                    credValue = textBoxPassword.Text + "||" + token;

                    crypto.SetSettings("ExcelPlugIn", "ManagementServer", comboBoxURL.Text);
                    crypto.SetSettings("ExcelPlugIn", "Username", comboBoxUsername.Text);
                    crypto.SetSettings("ExcelPlugIn", "Password", crypto.Encrypt(textBoxPassword.Text));
                    crypto.SetSettings("ExcelPlugIn", "Token", crypto.Encrypt(token));
                    crypto.SetSettings("ExcelPlugIn", "UseToken", checkBoxToken.Checked.ToString());

                    crypto.SetSettings("ExcelPlugIn\\Credentials", credName, crypto.Encrypt(credValue));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving settings to registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        #endregion

        private void comboBoxUsername_SelectedValueChanged(object sender, EventArgs e)
        {
            // MessageBox.Show("comboBoxUsername_SelectedValueChanged fired");
            RegistryKey Credentials = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\ExcelPlugIn\\Credentials");

            foreach (string valueName in Credentials.GetValueNames())
            {
                serverName = valueName.Substring(0, valueName.IndexOf("||"));
                userName = valueName.Substring(valueName.IndexOf("||") + 2);

                if (serverName == comboBoxURL.Text)
                {
                    // comboBoxUsername.Items.Add(userName);

                    if (userName == comboBoxUsername.Text)
                    {
                        if (userName == "Token")
                        {
                            string CredValue = crypto.Decrypt(Credentials.GetValue(valueName).ToString());
                            password = CredValue.Substring(0, CredValue.IndexOf("||"));
                            token = CredValue.Substring(CredValue.IndexOf("||") + 2);
                            textBoxPassword.Text = "";
                            textBoxToken.Text = token;

                            comboBoxUsername.Enabled = false;
                            textBoxPassword.Enabled = false;
                            textBoxToken.Enabled = true;
                            checkBoxToken.Checked = true;

                        }
                        else
                        {
                            string CredValue = crypto.Decrypt(Credentials.GetValue(valueName).ToString());
                            password = CredValue.Substring(0, CredValue.IndexOf("||"));
                            token = CredValue.Substring(CredValue.IndexOf("||") + 2);
                            textBoxPassword.Text = password;
                            textBoxToken.Text = "";

                            comboBoxUsername.Enabled = true;
                            textBoxPassword.Enabled = true;
                            textBoxToken.Enabled = false;
                            checkBoxToken.Checked = false;
                        }
                    }
                }
            }

        }

        private void comboBoxURL_SelectedValueChanged(object sender, EventArgs e)
        {
            RegistryKey Credentials = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\ExcelPlugIn\\Credentials");

            comboBoxUsername.Items.Clear();
            textBoxPassword.Text = "";

            foreach (string valueName in Credentials.GetValueNames())
            {
                serverName = valueName.Substring(0, valueName.IndexOf("||"));
                userName = valueName.Substring(valueName.IndexOf("||") + 2);

                if (serverName == comboBoxURL.Text)
                {
                    comboBoxUsername.Items.Add(userName);
                }
            }

            if (crypto.GetSettings("ExcelPlugIn", "ManagementServer") == comboBoxURL.Text)
            {
                comboBoxUsername.Text = crypto.GetSettings("ExcelPlugIn", "Username");
            }
            else
                comboBoxUsername.SelectedIndex = 0;
        }

        private void buttonClearOne_Click(object sender, EventArgs e)
        {
            try {
                RegistryKey Credentials = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\ExcelPlugIn");
                using (RegistryKey targetCred = Credentials.OpenSubKey("Credentials", true))
                {
                    string currentCred = comboBoxURL.Text + "||" + comboBoxUsername.Text;
                    // Delete the ID value.
                    targetCred.DeleteValue(currentCred);

                    if (comboBoxURL.Text == crypto.GetSettings("ExcelPlugIn", "ManagementServer"))
                    {
                        Credentials.DeleteValue("ManagementServer");
                        Credentials.DeleteValue("Username");
                        Credentials.DeleteValue("Password");
                        Credentials.DeleteValue("Token");
                    }
                    comboBoxURL.Text = "";
                    comboBoxUsername.Text = "";
                    textBoxPassword.Text = "";
                    textBoxToken.Text = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error clearing credential", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonClearAll_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure you want to clear all the credentials encrypted and remembered?", "Please confirm", MessageBoxButtons.YesNo);

                if (dialogResult == DialogResult.No)
                {
                    return;
                }

                RegistryKey Credentials = Registry.CurrentUser.CreateSubKey("Software\\SentinelOne\\ExcelPlugIn");

                Credentials.DeleteSubKeyTree("Credentials", false);

                Credentials.DeleteValue("ManagementServer", false);
                Credentials.DeleteValue("Username", false);
                Credentials.DeleteValue("Password", false);
                Credentials.DeleteValue("Token", false);

                comboBoxURL.Text = "";
                comboBoxUsername.Text = "";
                textBoxPassword.Text = "";
                textBoxToken.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error clearing all credentials", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void checkBoxToken_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxToken.Checked)
            {
                comboBoxUsername.Enabled = false;
                textBoxPassword.Enabled = false;
                textBoxToken.Enabled = true;
                comboBoxUsername.Text = "Token";
                textBoxPassword.Text = "";

                for (int i = 0; i < comboBoxUsername.Items.Count; i++)
                {
                    string value = comboBoxUsername.GetItemText(comboBoxUsername.Items[i]);
                    if (value == "Token")
                    {
                        comboBoxUsername.Text = value;
                        comboBoxUsername_SelectedValueChanged(comboBoxUsername, EventArgs.Empty);
                        break;
                    }
                }
            }
            else
            {
                comboBoxUsername.Enabled = true;
                textBoxPassword.Enabled = true;
                textBoxToken.Enabled = false;
                textBoxToken.Text = "";
                if (comboBoxUsername.Text == "Token")
                    comboBoxUsername.Text = "";

                for (int i = 0; i < comboBoxUsername.Items.Count; i++)
                {
                    string value = comboBoxUsername.GetItemText(comboBoxUsername.Items[i]);
                    if (value != "Token")
                    {
                        // comboBoxUsername.SelectedIndex = i;
                        comboBoxUsername.Text = value;
                        comboBoxUsername_SelectedValueChanged(comboBoxUsername, EventArgs.Empty);
                        break;
                    }
                }
            }
        }
    }

}
