using Common.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using Topshelf;
using System.Timers;
using System.IO;
using OpenPop.Mime;
using OpenPop.Pop3;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace SentinelOneNotificationService.Service
{
    class S1Notifier
    {
        private Timer workTimer = null;
        private string EmailSubject = "";
        private string EmailBody = "";

        private string ManagementServer = "";
        private string Token = "";

        private string Email_Pop_Server = "";
        private string Email_Pop_Port = "";
        private string Email_Pop_User = "";
        private string Email_Pop_Password = "";
        private string Email_Pop_SSL = "";

        private string Email_Smtp_Server = "";
        private string Email_Smtp_Port = "";
        private string Email_Smtp_User = "";
        private string Email_Smtp_Password = "";
        private string Email_Smtp_SSL = "";
        private string Email_Smtp_From = "";

        private string Email_To = "";

        public ILog Log { get; private set; }

        public S1Notifier(ILog logger)
        {

            // IocModule.cs needs to be updated in case new paramteres are added to this constructor


            if (logger == null)
                throw new ArgumentNullException(nameof(logger));

            Log = logger;

        }

        public bool Start(HostControl hostControl)
        {
            try
            {
                Log.Info($"Start command received.");

                var MyIni = new IniFile("settings.ini");

                ManagementServer = MyIni.GetValue("Server", "server");
                Token = MyIni.GetValue("Server", "token");

                Email_Pop_Server = MyIni.GetValue("POP", "email_pop_server");
                Email_Pop_Port = MyIni.GetValue("POP", "email_pop_port");
                Email_Pop_User = MyIni.GetValue("POP", "email_pop_user");
                Email_Pop_Password = MyIni.GetValue("POP", "email_pop_password");
                Email_Pop_SSL = MyIni.GetValue("POP", "email_pop_ssl");

                Email_Smtp_Server = MyIni.GetValue("SMTP", "email_smtp_server");
                Email_Smtp_Port = MyIni.GetValue("SMTP", "email_smtp_port");
                Email_Smtp_User = MyIni.GetValue("SMTP", "email_smtp_user");
                Email_Smtp_Password = MyIni.GetValue("SMTP", "email_smtp_password");
                Email_Smtp_SSL = MyIni.GetValue("SMTP", "email_smtp_ssl");
                Email_Smtp_From = MyIni.GetValue("SMTP", "email_smtp_from");

                Email_To = MyIni.GetValue("Recipients", "email_to");

                workTimer = new Timer();
                workTimer.Interval = 60000;
                workTimer.Elapsed += new System.Timers.ElapsedEventHandler(workTimer_Tick);
                workTimer.Enabled = true;
                workTimer_Tick(this, null);

                return true;
            }
            catch (Exception ex)
            {
                workTimer.Enabled = false;
                Log.Error($"Error starting service: " + ex.Message);
                return false;
            }
        }

        public bool Stop(HostControl hostControl)
        {

            Log.Trace($"Stop command received.");
            workTimer.Enabled = false;
            return true;

        }

        public bool Pause(HostControl hostControl)
        {

            Log.Trace($"Pause command received.");
            workTimer.Enabled = false;
            //TODO: Implement your service start routine.
            return true;

        }

        public bool Continue(HostControl hostControl)
        {

            Log.Trace($"Continue command received.");
            workTimer.Enabled = true;
            return true;

        }

        public bool Shutdown(HostControl hostControl)
        {

            Log.Trace($"Shutdown command received.");

            //TODO: Implement your service stop routine.
            return true;

        }

        private void workTimer_Tick(object sender, ElapsedEventArgs e)
        {
            try
            {
                // Log.Info($"Checking for new alerts...");

                string readEmailFile = "ReadEmails.id";

                List<String> seenMessages = new List<String>();
                if (File.Exists(readEmailFile))
                {
                    seenMessages = File.ReadAllLines(readEmailFile).ToList();
                }

                List<Message> NewMessages = FetchUnseenMessages(Email_Pop_Server,
                    Convert.ToInt32(Email_Pop_Port),
                    Convert.ToBoolean(Email_Pop_SSL),
                    Email_Pop_User,
                    Email_Pop_Password,
                    seenMessages);

                string threatid = "";

                for (int i = 0; i < NewMessages.Count; i++)

                {
                    Log.Info($"New email: " + NewMessages[i].Headers.Subject);

                    OpenPop.Mime.MessagePart htmlBody = NewMessages[i].FindFirstHtmlVersion();
                    EmailSubject = NewMessages[i].Headers.Subject;

                    if (htmlBody != null)
                    {
                        threatid = htmlBody.GetBodyAsText();
                    }

                    int start = threatid.IndexOf("analyze/threats/") + 16;
                    int end = threatid.IndexOf("/overview") - threatid.IndexOf("analyze/threats/") - 16;
                    threatid = threatid.Substring(start, end);
                    // Log.Info($"{nameof(Service.S1Notifier)} Threat ID: " + threatid);

                    var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                    {
                        NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                        MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                    };
                    var restClient = new RestClientInterface();
                    string resourceString = ManagementServer + "/web/api/v1.6/threats/" + threatid;
                    // threatid = "59d5239e75c5fa05cf6d74d0";
                    // 59d5239e75c5fa05cf6d74d0
                    // string resourceString = ManagementServer + "/web/api/v1.6/threats/59d5239e75c5fa05cf6d74d0";
                    restClient.EndPoint = resourceString;
                    restClient.Method = HttpVerb.GET;
                    var batch_string = restClient.MakeRequest(Token).ToString();
                    dynamic threat = Newtonsoft.Json.JsonConvert.DeserializeObject(batch_string, JsonSettings);
                    EmailBody = Newtonsoft.Json.Linq.JObject.Parse(batch_string.ToString()).ToString(Newtonsoft.Json.Formatting.Indented);
                    Log.Info($"Threat Details: " + EmailBody);

                    // Agent info ==============================================================================================
                    string agentid = threat.agent;
                    resourceString = ManagementServer + "/web/api/v1.6/agents/" + agentid;
                    restClient.EndPoint = resourceString;
                    restClient.Method = HttpVerb.GET;
                    var agent_string = restClient.MakeRequest(Token).ToString();
                    dynamic agent = Newtonsoft.Json.JsonConvert.DeserializeObject(agent_string, JsonSettings);

                    // Forensic info ==============================================================================================
                    resourceString = ManagementServer + "/web/api/v1.6/threats/" + threatid + "/forensics/export/json";
                    restClient.EndPoint = resourceString;
                    restClient.Method = HttpVerb.GET;
                    var forensic_string = restClient.MakeRequest(Token).ToString();
                    dynamic forensic = Newtonsoft.Json.JsonConvert.DeserializeObject(forensic_string, JsonSettings);
                    string forensic_output = Newtonsoft.Json.Linq.JObject.Parse(forensic_string.ToString()).ToString(Newtonsoft.Json.Formatting.Indented);

                    // Get the assembly information
                    System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
                    // Location is where the assembly is run from
                    string assemblyLocation = assemblyInfo.CodeBase;
                    assemblyLocation = assemblyLocation.Substring(0, assemblyLocation.LastIndexOf('/'));
                    string report_file = assemblyLocation + "/email.html";
                    string local_path = new Uri(report_file).LocalPath;
                    string report = File.ReadAllText(local_path);

                    string timestamp = threat.created_date.ToString();
                    string description = threat.description;
                    string resolved = threat.resolved;
                    string status = threat.mitigation_status;

                    string filename = threat.file_id.display_name;
                    string filepath = threat.file_id.path;
                    string filehash = threat.file_id.content_hash;

                    string rawdata = EmailBody;

                    switch (status)
                    {
                        case "0":
                            status = "<p class='S1Mitigated'>Mitigated</p>";
                            break;
                        case "1":
                            status = "<p class='S1Active'>Active</p>";
                            break;
                        case "2":
                            status = "<p class='S1Blocked'>Blocked</p>";
                            break;
                        case "3":
                            status = "<p class='S1Suspicious'>Suspicious</p>";
                            break;
                        case "4":
                            status = "<p class='S1Suspicious'>Pending</p>";
                            break;
                        case "5":
                            status = "<p class='S1Suspicious'>Suspicious cancelled</p>";
                            break;
                        default:
                            status = "<p class='S1Mitigated'>Mitigated</p>";
                            break;
                    }

                    switch (resolved.ToLower())
                    {
                        case "true":
                            resolved = "<p class='S1Green'>Yes</p>";
                            break;
                        case "false":
                            resolved = "<p class='S1Red'>No</p>";
                            break;
                        default:
                            resolved = "<p class='S1Red'>No</p>";
                            break;
                    }

                    string threatlink = "<a href='" + ManagementServer + "/#/analyze/threats/" + threatid + "/overview'>" + threatid + "</a>";
                    report = report.Replace("$threatid$", threatlink);
                    report = report.Replace("$timestamp$", timestamp);
                    report = report.Replace("$description$", description);
                    report = report.Replace("$resolved$", resolved);
                    report = report.Replace("$status$", status);

                    report = report.Replace("$filename$", filename);
                    report = report.Replace("$filepath$", filepath);
                    report = report.Replace("$filehash$", "<a href='https://www.virustotal.com/latest-scan/" + filehash + "'>" + filehash + "</a>");

                    string code = SyntaxHighlightJson(rawdata.Replace("\r\n", "<br/>").Replace(" ", "&nbsp;"));
                    // Log.Info($"{nameof(Service.S1Notifier)} Code: " + code);
                    report = report.Replace("$rawdata$", "<span class='Code'>" + code + "</span>");

                    string forensicdata = SyntaxHighlightJson(forensic_output.Replace("\r\n", "<br/>").Replace(" ", "&nbsp;"));
                    report = report.Replace("$forensicdata$", "<span class='Code'>" + forensicdata + "</span>");

                    string computername = agent.network_information.computer_name;
                    string os = agent.software_information.os_name;
                    string agentversion = agent.agent_version;
                    string agentip = agent.external_ip;
                    string machinetype = agent.hardware_information.machine_type;

                    report = report.Replace("$computername$", computername);
                    report = report.Replace("$os$", os);
                    report = report.Replace("$agentversion$", agentversion);
                    report = report.Replace("$agentip$", agentip);
                    report = report.Replace("$machinetype$", machinetype);

                    EmailBody = report;

                    SendEmail();
                }

                File.WriteAllLines(readEmailFile, seenMessages);
            }
            catch (Exception ex)
            {
                Log.Error($"Error in mail processing: " + ex.Message);
            }


        }

        /// <summary>
        /// Example showing:
        ///  - how to use UID's (unique ID's) of messages from the POP3 server
        ///  - how to download messages not seen before
        ///    (notice that the POP3 protocol cannot see if a message has been read on the server
        ///     before. Therefore the client need to maintain this state for itself)
        /// </summary>
        /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
        /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
        /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
        /// <param name="username">Username of the user on the server</param>
        /// <param name="password">Password of the user on the server</param>
        /// <param name="seenUids">
        /// List of UID's of all messages seen before.
        /// New message UID's will be added to the list.
        /// Consider using a HashSet if you are using >= 3.5 .NET
        /// </param>
        /// <returns>A List of new Messages on the server</returns>
        public static List<Message> FetchUnseenMessages(string hostname, int port, bool useSsl, string username, string password, List<string> seenUids)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Fetch all the current uids seen
                List<string> uids = client.GetMessageUids();

                // Create a list we can return with all new messages
                List<Message> newMessages = new List<Message>();

                // All the new messages not seen by the POP3 client
                for (int i = 0; i < uids.Count; i++)
                {
                    string currentUidOnServer = uids[i];
                    if (!seenUids.Contains(currentUidOnServer))
                    {
                        // We have not seen this message before.
                        // Download it and add this new uid to seen uids

                        // the uids list is in messageNumber order - meaning that the first
                        // uid in the list has messageNumber of 1, and the second has
                        // messageNumber 2. Therefore we can fetch the message using
                        // i + 1 since messageNumber should be in range [1, messageCount]
                        Message unseenMessage = client.GetMessage(i + 1);

                        // Add the message to the new messages
                        newMessages.Add(unseenMessage);

                        // Add the uid to the seen uids, as it has now been seen
                        seenUids.Add(currentUidOnServer);

                        // newReadIds.Add(currentUidOnServer);
                    }
                }

                // Return our new found messages
                return newMessages;
            }
        }


        private void SendEmail()
        {
            try
            {
                // MailAddress to = new MailAddress("leewei@sentinelone.com");
                string MailTo = Email_To;

                using (var mail = new MailMessage())
                {
                    mail.Subject = EmailSubject;
                    mail.IsBodyHtml = true;
                    mail.Body = EmailBody;
                    mail.From = new MailAddress(Email_Smtp_From);
                    string seperator = (MailTo.Contains(";") ? ";" : ",");
                    foreach (var address in MailTo.Split(new[] { seperator }, StringSplitOptions.RemoveEmptyEntries))
                    { mail.To.Add(address); }

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = Email_Smtp_Server;
                    smtp.Port = Convert.ToInt32(Email_Smtp_Port);
                    smtp.UseDefaultCredentials = Convert.ToBoolean(Email_Smtp_SSL);

                    smtp.Credentials = new NetworkCredential(Email_Smtp_User, Email_Smtp_Password);

                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }

                Log.Info($"Email sent");

            }
            catch (Exception ex)
            {
                Log.Error($"Error in sending email: " + ex.Message);
            }

        }

        public string SyntaxHighlightJson(string original)
        {
            return Regex.Replace(
              original,
              @"(""(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\""])*""(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)".Replace('\"', '"'),
              match => {
                  var cls = "number";
                  if (Regex.IsMatch(match.Value, @"^""".Replace('\"', '"')))
                  {
                      if (Regex.IsMatch(match.Value, ":$"))
                      {
                          cls = "key";
                      }
                      else {
                          cls = "string";
                      }
                  }
                  else if (Regex.IsMatch(match.Value, "true|false"))
                  {
                      cls = "boolean";
                  }
                  else if (Regex.IsMatch(match.Value, "null"))
                  {
                      cls = "null";
                  }
                  return "<span class=\"" + cls + "\">" + match + "</span>";
              });
        }
    }

}
