using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;

namespace S1ExcelPlugIn
{
    public partial class FormExecutiveSummary : Form
    {
        Crypto crypto = new Crypto();
        string userName = "";
        string mgmtServer = "";
        string serverName = "";
        string HTMLAttachment = "";
        string PDFAttachment = "";

        public FormExecutiveSummary()
        {
            InitializeComponent();
            labelEmailMessage.Visible = false;
            label1TestMessage.Visible = false;
            GetAll();

            mgmtServer = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
            serverName = mgmtServer.Substring(mgmtServer.IndexOf("//") + 2);

            if (textBoxSubject.Text.StartsWith("SentinelOne Server") &&
                textBoxSubject.Text.EndsWith("Summary Report Attached"))
            {
                textBoxSubject.Text = "SentinelOne Server (" + serverName + ") Summary Report Attached";
            }
        }

        private void Generate()
        {
            try
            {
                // webBrowserExecutiveSummary.DocumentText = "<html><body>Gathering data, but if the message is here for a minute, there is an error somewhere...<br/></body></html>";
                // webBrowserExecutiveSummary.ScriptErrorsSuppressed = true;
                labelEmailMessage.Visible = false;

                if (!checkBoxHTML.Checked && !checkBoxPDF.Checked)
                {
                    MessageBox.Show("Please select at least one of HTML or PDF output", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (checkBoxEmail.Checked && CheckEmailConfig() == false)
                    return;

                string token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                userName = crypto.GetSettings("ExcelPlugIn", "Username");

                // Agent Summary ==========================================================================================================
                string resourceString = mgmtServer + "/web/api/v1.6/agents/count-by-filters?participating_fields=" +
                    "software_information__os_type,software_information__os_arch,hardware_information__machine_type," +
                    "network_information__domain,network_status,configuration__learning_mode,is_pending_uninstall," +
                    "is_up_to_date,infected,policy_id,is_active";
                var restClient = new RestClientInterface(resourceString);
                restClient.EndPoint = resourceString;
                restClient.Method = HttpVerb.GET;
                var results = restClient.MakeRequest(token).ToString();
                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };
                dynamic agent_summary = Newtonsoft.Json.JsonConvert.DeserializeObject(results, JsonSettings);

                // webBrowserExecutiveSummary.DocumentText = "<html><body>" + results + "</body></html>";

                string AgentsInstalled = agent_summary.total_count.ToString();
                string AgentsIntalledFormatted = string.Format("{0:n0}", Convert.ToInt32(AgentsInstalled));

                string Connected = agent_summary.network_status.connected == null ? "0" : agent_summary.network_status.connected.ToString();
                string Disconnected = agent_summary.network_status.disconnected == null ? "0" : agent_summary.network_status.disconnected.ToString();
                int PercentConnected = (int)Math.Round((double)(100 * Convert.ToInt32(Connected)) / (Convert.ToInt32(Disconnected) + Convert.ToInt32(Connected)));

                string IsActiveTrue = agent_summary.is_active["true"] == null ? "0" : agent_summary.is_active["true"].ToString();
                string IsActiveFalse = agent_summary.is_active["false"] == null ? "0" : agent_summary.is_active["false"].ToString();
                // MessageBox.Show(IsActiveTrue + " - " + IsActiveFalse);
                int PercentActive = (int)Math.Round((double)(100 * Convert.ToInt32(IsActiveTrue)) / (Convert.ToInt32(IsActiveTrue) + Convert.ToInt32(IsActiveFalse)));

                // Get the assembly information
                System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
                // Location is where the assembly is run from 
                string assemblyLocation = assemblyInfo.CodeBase;
                assemblyLocation = assemblyLocation.Substring(0, assemblyLocation.LastIndexOf('/'));

                string timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                string reportPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                System.IO.Directory.CreateDirectory(reportPath + "\\SentinelOne Reports");

                string report_file = assemblyLocation + "/S1 Exec Report.html";
                string report_target = reportPath + "\\SentinelOne Reports\\" + "s1_" + serverName.Substring(0, serverName.IndexOf(".")) + "_" + timeStamp + ".html";

                string localPath = new Uri(report_file).LocalPath;
                string localTarget = new Uri(report_target).LocalPath;
                string workingdir = new Uri(assemblyLocation).LocalPath;

                string report = File.ReadAllText(localPath);

                #region Replace
                report = report.Replace("$Server$", serverName);

                report = report.Replace("$ReportPeriod$", Globals.ReportPeriod);
                report = report.Replace("$StartDate$", Globals.StartDate);
                report = report.Replace("$EndDate$", Globals.EndDate);

                report = report.Replace("$AI$", AgentsIntalledFormatted);
                report = report.Replace("$PA$", PercentActive.ToString() + "%");
                report = report.Replace("$MG$", Globals.MostAtRiskGroup);
                report = report.Replace("$MU$", Globals.MostAtRiskUser);
                report = report.Replace("$ME$", Globals.MostAtRiskEndpoint);

                report = report.Replace("$TABLE-THREATDATA$", Globals.ThreatData);
                report = report.Replace("$TABLE-DETECTIONENGINE$", Globals.DetectionEngine);
                report = report.Replace("$TABLE-INFECTEDFILES$", Globals.InfectedFiles);
                report = report.Replace("$TABLE-MOSTATRISKGROUPS$", Globals.MostAtRiskGroups);
                report = report.Replace("$TABLE-MOSTATRISKUSERS$", Globals.MostAtRiskUsers);
                report = report.Replace("$TABLE-MOSTATRISKENDPOINTS$", Globals.MostAtRiskEndpoints);

                report = report.Replace("$ThreatDataLabel$", Globals.ThreatDataLabel);
                report = report.Replace("$DetectionEnginesLabel$", Globals.DetectionEnginesLabel);
                report = report.Replace("$InfectedFilesLabel$", Globals.InfectedFilesLabel);
                report = report.Replace("$MostAtRiskGroupsLabel$", Globals.MostAtRiskGroupsLabel);
                report = report.Replace("$MostAtRiskUsersLabel$", Globals.MostAtRiskUsersLabel);
                report = report.Replace("$MostAtRiskEndpointsLabel$", Globals.MostAtRiskEndpointsLabel);

                report = report.Replace("$ThreatDataValue$", Globals.ThreatDataValue);
                report = report.Replace("$DetectionEnginesValue$", Globals.DetectionEnginesValue);
                report = report.Replace("$InfectedFilesValue$", Globals.InfectedFilesValue);
                report = report.Replace("$MostAtRiskGroupsValue$", Globals.MostAtRiskGroupsValue);
                report = report.Replace("$MostAtRiskUsersValue$", Globals.MostAtRiskUsersValue);
                report = report.Replace("$MostAtRiskEndpointsValue$", Globals.MostAtRiskEndpointsValue);

                report = report.Replace("$TABLE-NETWORKSTATUS$", Globals.NetworkStatus);
                report = report.Replace("$TABLE-ENDPOINTOS$", Globals.EndpointOS);
                report = report.Replace("$TABLE-ENDPOINTVERSION$", Globals.EndpointVersion);

                report = report.Replace("$NetworkStatusLabel$", Globals.NetworkStatusLabel);
                report = report.Replace("$EndpointOSLabel$", Globals.EndpointOSLabel);
                report = report.Replace("$EndpointVersionLabel$", Globals.EndpointVersionLabel);

                report = report.Replace("$NetworkStatusValue$", Globals.NetworkStatusValue);
                report = report.Replace("$EndpointOSValue$", Globals.EndpointOSValue);
                report = report.Replace("$EndpointVersionValue$", Globals.EndpointVersionValue);

                report = report.Replace("$TABLE-TOPAPPLICATIONS$", Globals.TopApplications);
                report = report.Replace("$TopApplicationsLabel$", Globals.TopApplicationsLabel);
                report = report.Replace("$TopApplicationsValue$", Globals.TopApplicationsValue);

                report = report.Replace("$ReportTimestamp$", DateTime.UtcNow.ToString("f") + " (UTC)");
                report = report.Replace("$ReportUser$", userName);
                #endregion

                // Threat Summary ================================================================================================================
                resourceString = mgmtServer + "/web/api/v1.6/threats/summary";
                restClient = new RestClientInterface(resourceString);
                restClient.EndPoint = resourceString;
                restClient.Method = HttpVerb.GET;
                var resultsThreats = restClient.MakeRequest(token).ToString();
                dynamic threat_summary = Newtonsoft.Json.JsonConvert.DeserializeObject(resultsThreats, JsonSettings);

                string ActiveThreats = threat_summary.active.ToString();
                string MitigatedThreats = threat_summary.mitigated.ToString();
                string SuspiciousThreats = threat_summary.suspicious.ToString();
                string BloackedThreats = threat_summary.blocked.ToString();

                report = report.Replace("$TT$", string.Format("{0:n0}", Convert.ToInt32(Globals.TotalThreats)));
                if (Globals.ActiveThreats == "") Globals.ActiveThreats = "0";
                report = report.Replace("$AT$", string.Format("{0:n0}", Convert.ToInt32(Globals.ActiveThreats)));
                report = report.Replace("$AUT$", string.Format("{0:n0}", Convert.ToInt32(Globals.ActiveAndUnresolvedThreats)));
                report = report.Replace("$UnresolvedThreatOnly$", Globals.UnresolvedThreatOnly.ToString().ToLower());

                File.WriteAllText(localTarget, report);

                HTMLAttachment = localTarget;

                if (checkBoxHTML.Checked)
                {
                    System.Diagnostics.Process.Start(localTarget);
                }

                // webBrowserExecutiveSummary.Url = new Uri("file://127.0.0.1/c$/temp/S1Report.html");
                // webBrowserExecutiveSummary.IsWebBrowserContextMenuEnabled = false;
                // webBrowserExecutiveSummary.Url = new Uri(report_target);

                if (checkBoxPDF.Checked)
                {
                    Process p = new Process();
                    p.StartInfo.WorkingDirectory = workingdir;
                    p.StartInfo.FileName = workingdir + "\\wkhtmltopdf.exe";
                    p.StartInfo.Arguments = "--lowquality " +
                                            "--print-media-type " +
                                            "--page-size Letter " +
                                            "--footer-spacing 2 " +
                                            "--footer-right \"[page] of [toPage]\" " +
                                            "--footer-left \"[isodate]    [time]\" " +
                                            "--footer-font-size 6 " +
                                            "\"" + localTarget + "\"" +
                                            " \"" + Path.ChangeExtension(localTarget, ".pdf") + "\"";

                    // Stop the process from opening a new window
                    p.StartInfo.CreateNoWindow = true;

                    p.StartInfo.UseShellExecute = false;
                    p.StartInfo.RedirectStandardOutput = true;
                    p.StartInfo.RedirectStandardError = true;
                    p.Start();
                    string output = p.StandardError.ReadToEnd();
                    p.WaitForExit();
                    p.Close();

                    PDFAttachment = Path.ChangeExtension(localTarget, ".pdf");
                    System.Diagnostics.Process.Start(Path.ChangeExtension(localTarget, ".pdf"));
                }

                if (checkBoxEmail.Checked)
                {
                    SendEmail();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error extracting report data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
            Generate();
        }

        private void SendEmail()
        {
            try
            {
                buttonReport.Enabled = false;
                buttonReport.Update();

                MailAddress to = new MailAddress(textBoxMailTo.Text);

                using (var mail = new MailMessage())
                {
                    mail.Subject = textBoxSubject.Text;

                    mail.Body = textBoxBody.Text + "\r\n\r\n";
                    mail.From = new MailAddress(textBoxMailFrom.Text);
                    string seperator = (textBoxMailTo.Text.Contains(";") ? ";" : ",");
                    foreach (var address in textBoxMailTo.Text.Split(new[] { seperator }, StringSplitOptions.RemoveEmptyEntries))
                    { mail.To.Add(address); }

                    System.Net.Mail.Attachment mail_attachment1;
                    if (checkBoxHTML.Checked)
                    {
                        mail_attachment1 = new System.Net.Mail.Attachment(HTMLAttachment);
                        mail.Attachments.Add(mail_attachment1);
                    }

                    System.Net.Mail.Attachment mail_attachment2;
                    if (checkBoxPDF.Checked)
                    {
                        mail_attachment2 = new System.Net.Mail.Attachment(PDFAttachment);
                        mail.Attachments.Add(mail_attachment2);
                    }

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = textBoxSMTPHost.Text;
                    smtp.Port = Convert.ToInt32(textBoxSMTPPort.Text);
                    smtp.UseDefaultCredentials = false;

                    if (textBoxUsername.Text != null)
                    {
                        smtp.Credentials = new NetworkCredential(textBoxUsername.Text, textBoxPassword.Text);
                    }

                    smtp.EnableSsl = checkBoxSSL.Checked;
                    smtp.Send(mail);

                    labelEmailMessage.Visible = true;
                    buttonReport.Enabled = true;
                    buttonReport.Update();

                }
            }
            catch (Exception ex)
            {
                buttonReport.Enabled = true;
                buttonReport.Update();
                MessageBox.Show(ex.Message, "Error sending email", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private bool CheckEmailConfig()
        {
            if (textBoxSMTPHost.Text == "" ||
                textBoxSMTPPort.Text == "" ||
                textBoxUsername.Text == "" ||
                textBoxPassword.Text == "" ||
                textBoxMailFrom.Text == "")
            {
                MessageBox.Show("Email server information not configured", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            else
                return true;
        }

        private void checkBoxEmail_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxEmail.Checked)
            {
                textBoxMailTo.Enabled = true;
                textBoxSubject.Enabled = true;
                textBoxBody.Enabled = true;
            }
            else
            {
                textBoxMailTo.Enabled = false;
                textBoxSubject.Enabled = false;
                textBoxBody.Enabled = false;
            }
        }

        #region Saving and Retrieving from Registry

        private void SaveAll()
        {
            try
            {
                crypto.SetSettings("ExcelPlugIn", "GenerateHTML", checkBoxHTML.Checked.ToString());
                crypto.SetSettings("ExcelPlugIn", "GeneratePDF", checkBoxPDF.Checked.ToString());
                crypto.SetSettings("ExcelPlugIn", "SendEmail", checkBoxEmail.Checked.ToString());

                crypto.SetSettings("ExcelPlugIn", "EmailMailTo", textBoxMailTo.Text);
                crypto.SetSettings("ExcelPlugIn", "EmailSubject", textBoxSubject.Text);
                crypto.SetSettings("ExcelPlugIn", "EmailBody", textBoxBody.Text);

                crypto.SetSettings("ExcelPlugIn", "EmailSMTPHost", textBoxSMTPHost.Text);
                crypto.SetSettings("ExcelPlugIn", "EmailSMTPPort", textBoxSMTPPort.Text);
                crypto.SetSettings("ExcelPlugIn", "EmailEnableSSL", checkBoxSSL.Checked.ToString());
                crypto.SetSettings("ExcelPlugIn", "EmailUsername", textBoxUsername.Text);
                crypto.SetSettings("ExcelPlugIn", "EmailPassword", crypto.Encrypt(textBoxPassword.Text));
                crypto.SetSettings("ExcelPlugIn", "EmailMailFrom", textBoxMailFrom.Text);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving settings to registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void GetAll()
        {
            try
            {
                if (crypto.GetSettings("ExcelPlugIn", "GenerateHTML") == null)
                    checkBoxHTML.Checked = true;
                else
                    checkBoxHTML.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "GenerateHTML"));

                if (crypto.GetSettings("ExcelPlugIn", "GeneratePDF") == null)
                    checkBoxPDF.Checked = true;
                else
                    checkBoxPDF.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "GeneratePDF"));

                if (crypto.GetSettings("ExcelPlugIn", "SendEmail") == null)
                    checkBoxEmail.Checked = true;
                else
                    checkBoxEmail.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "SendEmail"));

                textBoxMailTo.Text = crypto.GetSettings("ExcelPlugIn", "EmailMailTo");

                if (crypto.GetSettings("ExcelPlugIn", "EmailSubject") == null)
                    textBoxSubject.Text = "SentinelOne Server Summary Report Attached";
                else
                    textBoxSubject.Text = crypto.GetSettings("ExcelPlugIn", "EmailSubject");

                if (crypto.GetSettings("ExcelPlugIn", "EmailBody") == null)
                    textBoxBody.Text = "Please see attached report";
                else
                    textBoxBody.Text = crypto.GetSettings("ExcelPlugIn", "EmailBody");

                if (crypto.GetSettings("ExcelPlugIn", "EmailSMTPHost") == null)
                    textBoxSMTPHost.Text = "smtp.gmail.com";
                else
                    textBoxSMTPHost.Text = crypto.GetSettings("ExcelPlugIn", "EmailSMTPHost");

                if (crypto.GetSettings("ExcelPlugIn", "EmailSMTPPort") == null)
                    textBoxSMTPPort.Text = "587";
                else
                    textBoxSMTPPort.Text = crypto.GetSettings("ExcelPlugIn", "EmailSMTPPort");

                if (crypto.GetSettings("ExcelPlugIn", "EmailEnableSSL") == null)
                    checkBoxSSL.Checked = true;
                else
                    checkBoxSSL.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "EmailEnableSSL"));

                if (crypto.GetSettings("ExcelPlugIn", "EmailUsername") == null)
                    textBoxUsername.Text = "sentinelonereport@gmail.com";
                else
                    textBoxUsername.Text = crypto.GetSettings("ExcelPlugIn", "EmailUsername");

                if (crypto.GetSettings("ExcelPlugIn", "EmailPassword") == null)
                    textBoxPassword.Text = "idcafkjdumnicedv";
                else
                    textBoxPassword.Text = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "EmailPassword"));

                if (crypto.GetSettings("ExcelPlugIn", "EmailMailFrom") == null)
                    textBoxMailFrom.Text = "sentinelonereport@gmail.com";
                else
                    textBoxMailFrom.Text = crypto.GetSettings("ExcelPlugIn", "EmailMailFrom");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error saving settings to registry", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        #endregion

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormExecutiveSummary_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveAll();
        }

        private void buttonTest_Click(object sender, EventArgs e)
        {
            try
            {
                labelEmailMessage.Visible = false;
                label1TestMessage.Update();
                buttonTest.Enabled = false;
                buttonTest.Update();

                using (var mail = new MailMessage())
                {
                    mail.Subject = "SentinelOne Summary Report test email";
                    mail.Body = "SentinelOne Summary Report test email" + "\r\n\r\n";
                    mail.From = new MailAddress(textBoxMailFrom.Text);
                    mail.To.Add(new MailAddress(textBoxMailFrom.Text));

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = textBoxSMTPHost.Text;
                    smtp.Port = Convert.ToInt32(textBoxSMTPPort.Text);
                    smtp.UseDefaultCredentials = false;

                    if (textBoxUsername.Text != null)
                    {
                        smtp.Credentials = new NetworkCredential(textBoxUsername.Text, textBoxPassword.Text);
                    }

                    smtp.EnableSsl = checkBoxSSL.Checked;
                    smtp.Send(mail);

                    SaveAll();

                    label1TestMessage.Visible = true;
                    buttonTest.Enabled = true;
                    buttonTest.Update();
                }
            }
            catch (Exception ex)
            {
                buttonTest.Enabled = true;
                label1TestMessage.Visible = false;
                label1TestMessage.Update();
                if (ex.Message.Contains("5.5.1 Authentication Required") && textBoxSMTPHost.Text.ToLower().Contains("gmail"))
                {
                    MessageBox.Show(ex.Message + "\r\n\r\n" + "If you are using Gmail, the solution might be to use the feature \r\n\"Sign in using App Passwords\"\r\nhttps://support.google.com/accounts/answer/185833?hl=en" , "Error sending test email", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(ex.Message, "Error sending test email", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        }


    }
}
