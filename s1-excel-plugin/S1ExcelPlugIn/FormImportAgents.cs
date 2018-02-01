using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Reflection;
using Newtonsoft.Json.Linq;
using AddinExpress.MSO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Linq;
using System.Collections;
using System.Globalization;
using System.Diagnostics;
using System.Text;

namespace S1ExcelPlugIn
{
    public partial class FormImportAgents : Form
    {
        #region Class variables
        Crypto crypto = new Crypto();
        string userName = "";
        string mgmtServer = "";

        int colCount = 0;
        int rowCount = 0;

        int r_color = 106;
        int g_color = 34;
        int b_color = 182;

        // Color for reference column light purple
        int r_color_ref = 233;
        int g_color_ref = 227;
        int b_color_ref = 244;

        // Color for reference column header medium purple
        int r_color_refh = 221;
        int g_color_refh = 211;
        int b_color_refh = 238;

        ArrayList dateTimeCol = new ArrayList();
        ArrayList integerCol = new ArrayList();
        ArrayList referenceCol = new ArrayList();

        Stopwatch stopWatch = new Stopwatch();
        ExcelHelper eHelper = new ExcelHelper();

        FormMessages formMsg = new FormMessages();

        string resourceString = "";

        #endregion

        public FormImportAgents()
        {
            InitializeComponent();

            try
            {
                if (crypto.GetSettings("ExcelPlugIn", "ManagementServer") == null || crypto.GetSettings("ExcelPlugIn", "Token") == null)
                {
                    MessageBox.Show("Please log into the SentinelOne Management Server first", "Please login", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Load += (s, e) => Close();
                    return;
                }

                // Restoring previously specified agent import parameters

                if (crypto.GetSettings("ExcelPlugIn", "AgentReturnAll") == null)
                    checkBoxAll.Checked = true;
                else
                    checkBoxAll.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "AgentReturnAll"));

                if (crypto.GetSettings("ExcelPlugIn", "AgentQuery") == null)
                    textBoxQuery.Text = "";
                else
                    textBoxQuery.Text = crypto.GetSettings("ExcelPlugIn", "AgentQuery");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error initialzing saved parameters, no harm done", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Button Events
        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonImportData_Click(object sender, EventArgs e)
        {
            GetData();
            Close();
        }

        private void buttonImportAndGenerate_Click(object sender, EventArgs e)
        {

            GetData();


            Close();
        }

        #endregion 

        public void GetData()
        {
            try
            {
                #region Create worksheet
                // Creates a worksheet for "Threats" if one does not already exist
                // =================================================================================================================
                Excel.Worksheet threatsWorksheet;
                try
                {
                    threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Agent Data");
                    threatsWorksheet.Activate();
                }
                catch
                {
                    threatsWorksheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    threatsWorksheet.Name = "Agent Data";
                    threatsWorksheet.Activate();
                }
                #endregion

                #region Clear worksheet
                // Clear spreadsheet
                eHelper.Clear();
                #endregion

                #region Get Data

                #region Get data from server
                // Extract data from SentinelOne Management Server (Looping Calls)
                // =================================================================================================================
                mgmtServer = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                string token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                userName = crypto.GetSettings("ExcelPlugIn", "Username");

                crypto.SetSettings("ExcelPlugIn", "AgentQuery", textBoxQuery.Text);
                crypto.SetSettings("ExcelPlugIn", "AgentReturnAll", checkBoxAll.Checked.ToString());

                string query = "";
                if (textBoxQuery.Text != "" && checkBoxAll.Checked == false)
                {
                    query = "&query=" + textBoxQuery.Text;
                }
                else
                {
                }

                string limit = "limit=" + Globals.ApiBatch.ToString();
                bool Gogo = true;
                string skip = "&skip=";
                string last_id = "";
                int skip_count = 0;
                StringBuilder results = new StringBuilder("[");
                int maxColumnWidth = 80;
                dynamic threats = "";
                dynamic agents = "";
                dynamic agentsIterate = "";
                int rowCountTemp = 0;

                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };

                formMsg.StopProcessing = false;
                formMsg.Message("Loading agent data: " + rowCount.ToString("N0"),
                        eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);

                this.Opacity = 0;
                this.ShowInTaskbar = false;

                int percentComplete = 0;

                var restClient = new RestClientInterface();

                stopWatch.Start();

                if (Globals.ServerVersion.Contains("v1."))
                {
                    #region v1 API using limit
                    while (Gogo)
                    {
                        resourceString = mgmtServer + "/web/api/v1.6/agents?" + limit + skip + skip_count.ToString() + query;
                        restClient.EndPoint = resourceString;
                        restClient.Method = HttpVerb.GET;
                        var batch_string = restClient.MakeRequest(token).ToString();
                        threats = Newtonsoft.Json.JsonConvert.DeserializeObject(batch_string, JsonSettings);
                        rowCountTemp = (int)threats.Count;
                        skip_count = skip_count + Globals.ApiBatch;
                        results.Append(batch_string.TrimStart('[').TrimEnd(']', '\r', '\n')).Append(",");

                        rowCount = rowCount + rowCountTemp;

                        if (rowCountTemp < Globals.ApiBatch)
                            Gogo = false;

                        percentComplete = (int)Math.Round((double)(100 * rowCount) / Globals.TotalAgents);
                        formMsg.Message("Loading " + rowCount.ToString("N0") + " of " + Globals.TotalAgents.ToString("N0") + " agents (" + percentComplete.ToString() + "%)...",
                            eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);
                        if (formMsg.StopProcessing == true)
                        {
                            formMsg.Hide();
                            return;
                        }
                    }
                    #endregion
                }
                else
                {
                    #region v2 API using iterator
                    while (Gogo)
                    {
                        resourceString = mgmtServer + "/web/api/v1.6/agents/iterator?" + limit + last_id + query;
                        restClient.EndPoint = resourceString;
                        restClient.Method = HttpVerb.GET;
                        var res = restClient.MakeRequest(token).ToString();
                        agentsIterate = Newtonsoft.Json.JsonConvert.DeserializeObject(res, JsonSettings);
                        rowCountTemp = (int)agentsIterate.data.Count;
                        agents = agentsIterate.data;
                        last_id = "&last_id=" + agentsIterate.last_id;
                        skip_count = skip_count + Globals.ApiBatch;

                        results.Append(agents.ToString().TrimStart('[').TrimEnd(']', '\r', '\n')).Append(",");

                        rowCount = rowCount + rowCountTemp;

                        if (agentsIterate.last_id == null)
                            Gogo = false;

                        percentComplete = (int)Math.Round((double)(100 * rowCount) / Globals.TotalAgents);
                        formMsg.Message("Iterating data for " + rowCount.ToString("N0") + " of " + Globals.TotalAgents.ToString("N0") + " agents for reference (" + percentComplete.ToString() + "%)...",
                               eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);
                        if (formMsg.StopProcessing == true)
                        {
                            formMsg.Hide();
                            return;
                        }
                    }
                    #endregion
                }

                stopWatch.Stop();
                formMsg.Hide();

                Globals.ApiUrl = resourceString;

                // results.ToString().TrimEnd(',', ' ');
                results.Length--;
                results.Append("]");

                threats = Newtonsoft.Json.JsonConvert.DeserializeObject(results.ToString(), JsonSettings);

                #endregion

                #region Parse attribute headers
                JArray ja = JArray.Parse(results.ToString());
                // Stop processing if no data found
                if (ja.Count == 0)
                {
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "No agent data found";
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Select();
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Columns.AutoFit();
                    formMsg.Hide();
                    return;
                }

                Dictionary<string, object> dictObj = ja[0].ToObject<Dictionary<string, object>>();
                List<KeyValuePair<string, string>> AttribCollection = new List<KeyValuePair<string, string>>();
                int AttribNo = 0;

                for (int index = 0; index < dictObj.Count; index++)
                {
                    var Level1 = dictObj.ElementAt(index);
                    string dataType = "String";
                    dataType = eHelper.GetDataType(Level1);

                    if (dataType == "Object")
                    {
                        foreach (KeyValuePair<string, object> Level2 in ((JObject)Level1.Value).ToObject<Dictionary<string, object>>())
                        {
                            dataType = eHelper.GetDataType(Level2);
                            if (dataType == "Object")
                            {
                                foreach (KeyValuePair<string, object> Level3 in ((JObject)Level2.Value).ToObject<Dictionary<string, object>>())
                                {
                                    dataType = eHelper.GetDataType(Level3);
                                    AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + Level2.Key + "." + Level3.Key, dataType));
                                    AttribNo++;
                                    AddFormatter(AttribNo-1, dataType, Level1.Key + "." + Level2.Key + "." + Level3.Key);
                                }
                            }
                            else if (dataType == "Array" && Level1.Key == "network_information")
                            {
                                AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + Level2.Key, dataType));
                                AttribNo++;

                                AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "ip_addresses", "String"));
                                AttribNo++;
                                // AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "ip_address_2", "String"));
                                // AttribNo++;
                                // AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "ip_address_3", "String"));
                                // AttribNo++;
                            }
                            else
                            {
                                AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + Level2.Key, dataType));
                                AttribNo++;
                                AddFormatter(AttribNo - 1, dataType, Level1.Key + "." + Level2.Key);
                            }
                        }
                    }

                    else if (dataType == "Array" && Level1.Key == "engine_data")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "engine", "String"));
                        AttribNo++;
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "asset_name", "String"));
                        AttribNo++;
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "asset_version", "String"));
                        AttribNo++;
                    }
                    else if (Level1.Key == "group_id")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("group_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "group_name");
                    }
                    else
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo-1, dataType, Level1.Key);
                    }
                }
                colCount = AttribCollection.Count;
                #endregion

                #endregion

                #region Create Headings
                // Create headings
                // =================================================================================================================
                Excel.Range titleRow = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", eHelper.ExcelColumnLetter(colCount-1) + "1");
                titleRow.Select();
                titleRow.RowHeight = 33;
                titleRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color));

                Excel.Range rowSeparator = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A3", eHelper.ExcelColumnLetter(colCount - 1) + "3");
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color)); // 
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1; // xlContinuous
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = 4; // xlThick
                #endregion

                #region Write data
                // Write all the data rows
                // =================================================================================================================
                string[,] dataBlock = new string[rowCount, colCount];
                dynamic temp = "";

                string[] prop;
                string header0 = "";
                string[] ip_interfaces;
                bool byPass2 = false;
                bool byPass3 = false;

                Excel.Worksheet lookupSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                int rows = Convert.ToInt32(lookupSheet.Cells[3, 25].Value) + 4;

                formMsg.Message(rowCount.ToString("N0") + " agent data loaded, now processing locally in Excel...", "", allowCancel: false);

                for (int i = 0; i < rowCount; i++)
                {
                    byPass2 = false;
                    byPass3 = false;

                    for (int j = 0; j < AttribCollection.Count; j++)
                    {
                        prop = AttribCollection[j].Key.Split('.');
                        header0 = prop[0];

                        if (prop.Length == 1)
                            temp = (threats[i][prop[0]] == null) ? "null" : threats[i][prop[0]].ToString();
                        else if (prop.Length == 2)
                        {
                            if (prop[1] == "ip_addresses")
                            {
                                JArray ipa = JArray.Parse(threats[i][prop[0]]["interfaces"].ToString());
                                int ipa_count = ipa.Count;
                                temp = "";

                                for (int k = 0; k < ipa.Count; k++)
                                {
                                    if (ipa[k]["physical"].ToString() != "00:00:00:00:00:00" && ipa[k]["inet"].ToObject<string[]>().Length > 0)
                                    {
                                        temp = temp + ipa[k]["inet"].ToObject<string[]>()[0] + "; ";
                                    }
                                }

                                temp = temp.ToString().TrimEnd(';', ' ');
                                /*
                                if (prop[1] == "ip_address_1" && ipa_count > 0)
                                {
                                    if (ipa[0]["physical"].ToString() != "00:00:00:00:00:00")
                                    {
                                        ip_interfaces = ipa[0]["inet"].ToObject<string[]>();
                                        temp = ip_interfaces.Length > 0 ? ip_interfaces[0] : "";
                                    }
                                    else if (ipa_count > 1 && ipa[1]["physical"].ToString() != "00:00:00:00:00:00")
                                    {
                                        ip_interfaces = ipa[1]["inet"].ToObject<string[]>();
                                        temp = ip_interfaces.Length > 0 ? ip_interfaces[0] : "";
                                        byPass2 = true;
                                    }
                                    else if (ipa_count > 2 && ipa[2]["physical"].ToString() != "00:00:00:00:00:00")
                                    {
                                        ip_interfaces = ipa[2]["inet"].ToObject<string[]>();
                                        temp = ip_interfaces.Length > 0 ? ip_interfaces[0] : "";
                                        byPass3 = true;
                                    }
                                }
                                else if (prop[1] == "ip_address_2" && ipa_count > 1)
                                {
                                    if (ipa[1]["physical"].ToString() != "00:00:00:00:00:00" && byPass2 == false)
                                    {
                                        ip_interfaces = ipa[1]["inet"].ToObject<string[]>();
                                        temp = ip_interfaces.Length > 0 ? ip_interfaces[0] : "";
                                    }
                                    else if (ipa_count > 2 && ipa[2]["physical"].ToString() != "00:00:00:00:00:00" && byPass3 == false)
                                    {
                                        ip_interfaces = ipa[2]["inet"].ToObject<string[]>();
                                        temp = ip_interfaces.Length > 0 ? ip_interfaces[0] : "";
                                        byPass3 = true;
                                    }
                                }
                                else if (prop[1] == "ip_address_3" && ipa_count > 2)
                                {
                                    if (ipa[2]["physical"].ToString() != "00:00:00:00:00:00" && byPass3 == false)
                                    {
                                        ip_interfaces = ipa[2]["inet"].ToObject<string[]>();
                                        temp = ip_interfaces.Length > 0 ? ip_interfaces[0] : "";
                                    }
                                }
                                */

                            }
                            else
                            {
                                temp = (threats[i][prop[0]][prop[1]] == null) ? "null" : threats[i][prop[0]][prop[1]].ToString();
                            }
                        }
                        else if (prop.Length == 3)
                            temp = (threats[i][prop[0]][prop[1]][prop[2]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]].ToString();
                        else if (prop.Length == 4)
                            temp = (threats[i][prop[0]][prop[1]][prop[2]][prop[3]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]][prop[3]].ToString();

                        if (prop.Length == 2 && prop[1] == "os_start_time" )
                        {
                            DateTime d2 = DateTime.Parse(temp, null, System.Globalization.DateTimeStyles.RoundtripKind);
                            temp = "=DATE(" + d2.ToString("yyyy,MM,dd") + ")+TIME(" + d2.ToString("HH,mm,ss") + ")";
                        }

                        switch (header0)
                        {
                            #region Lookups
                            case "group_name":
                                {
                                    temp = "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(j - 1) + (i + 5).ToString() + ",'Lookup Tables'!Y4:Z" + rows.ToString() + ",2,FALSE),\"null\")";
                                    break;
                                }

                            case "created_date":
                            case "last_active_date":
                            case "meta_data":
                                {
                                    if (temp != "null")
                                    {
                                        DateTime d2 = DateTime.Parse(temp, null, System.Globalization.DateTimeStyles.RoundtripKind);
                                        temp = "=DATE(" + d2.ToString("yyyy,MM,dd") + ")+TIME(" + d2.ToString("HH,mm,ss") + ")";
                                    }
                                    else
                                    {
                                        temp = "";
                                    }
                                    break;
                                }
                            #endregion

                        }

                        dataBlock[i, j] = temp;
                        }
                    }

                Excel.Range range;
                range = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A5", Missing.Value);

                if (threats.Count > 0)
                {
                    range = range.Resize[rowCount, colCount];
                    range.Cells.ClearFormats();
                    // Writes the array into Excel
                    // This is probably the single thing that sped up the report the most, by writing array
                    range.Value = dataBlock;
                    range.Font.Size = "10";
                    range.RowHeight = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Height;
                }
                #endregion

                #region Column Headers
                // Column Headers
                // =================================================================================================================
                eHelper.WriteHeaders("Agent Data", colCount, rowCount, stopWatch);

                // This block writes the column headers
                string hd = "";
                for (int m = 0; m < AttribCollection.Count; m++)
                {
                    TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
                    hd = textInfo.ToTitleCase(AttribCollection[m].Key).Replace('_', ' ');
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[4, m + 1] = hd;
                }
                #endregion

                #region Formatting
                // Create the data filter
                // =================================================================================================================
                Excel.Range range2;
                range2 = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A4", Missing.Value);
                range2 = range2.Resize[rowCount + 1, colCount];
                range2.Select();
                // range2.AutoFilter("1", "<>", Excel.XlAutoFilterOperator.xlOr, "", true);
                range2.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlOr, Type.Missing, true);

                // Format number columns 
                // =================================================================================================================
                for (int i = 0; i < integerCol.Count; i++)
                {
                    Excel.Range rangeInt;
                    rangeInt = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(eHelper.ExcelColumnLetter((int)integerCol[i]) + "5", eHelper.ExcelColumnLetter((int)integerCol[i]) + (rowCount + 5).ToString());
                    rangeInt.NumberFormat = "0";
                    rangeInt.Cells.HorizontalAlignment = -4152; // Right align the number
                    rangeInt.Value = rangeInt.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                }

                // Format date time columns
                // =================================================================================================================
                for (int j = 0; j < dateTimeCol.Count; j++)
                {
                    Excel.Range rangeTime;
                    rangeTime = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(eHelper.ExcelColumnLetter((int)dateTimeCol[j]) + "5", eHelper.ExcelColumnLetter((int)dateTimeCol[j]) + (rowCount + 5).ToString());
                    rangeTime.NumberFormat = "[$-409]yyyy/mm/dd hh:mm AM/PM;@";
                    rangeTime.Cells.HorizontalAlignment = -4152; // Right align the Time
                    rangeTime.Value = rangeTime.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                }

                // Format reference columns
                // =================================================================================================================
                for (int k = 0; k < referenceCol.Count; k++)
                {
                    Excel.Range rangeReference;
                    rangeReference = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(eHelper.ExcelColumnLetter((int)referenceCol[k]) + "5", eHelper.ExcelColumnLetter((int)referenceCol[k]) + (rowCount + 4).ToString());
                    rangeReference.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color_ref, g_color_ref, b_color_ref));
                    rangeReference.Formula = rangeReference.Value;

                    rangeReference = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(eHelper.ExcelColumnLetter((int)referenceCol[k]) + "4", eHelper.ExcelColumnLetter((int)referenceCol[k]) + "4");
                    rangeReference.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color_refh, g_color_refh, b_color_refh));
                    // rangeReference.Value = rangeReference.Value;
                }

                if (referenceCol.Count > 0)
                {
                    Excel.Range rangeReference;
                    rangeReference = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("E2", "E2");
                    rangeReference.Value = "Reference lookup data";
                    rangeReference.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color_refh, g_color_refh, b_color_refh));
                }


                // Autofit column width
                // Formats the column width nicely to see the content, but not too wide and limit to maximum of 80
                // =================================================================================================================
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Columns.AutoFit();
                for (int i = 0; i < colCount; i++)
                {
                    Excel.Range rangeCol;
                    rangeCol = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(eHelper.ExcelColumnLetter(i) + "1", eHelper.ExcelColumnLetter(i) + "1");
                    if (Convert.ToInt32(rangeCol.ColumnWidth) > maxColumnWidth)
                    {
                        rangeCol.ColumnWidth = maxColumnWidth;
                    }
                }

                // Place the cursor in cell A2 - which is at the start of the document
                // =================================================================================================================
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Select();
                #endregion

                formMsg.Message("Creating pivot tables and charts for the agent data...", "", allowCancel: false);

                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = false;
                GenerateReport();
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = true;

                formMsg.Hide();

                // This is saved for viewing in the Show API window
                Globals.ApiResults = JArray.Parse(results.ToString()).ToString();

            }
            catch (Exception ex)
            {
                formMsg.Hide();
                MessageBox.Show(ex.Message, "Error extracting Agent data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GenerateReport()
        {
            #region Initialize
            Excel.Workbook activeWorkBook = null;
            Excel.Worksheet pivotWorkSheet = null;
            #endregion

            try
            {
                activeWorkBook = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook;
                try
                {
                    pivotWorkSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Agent Reports");
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet.Application.DisplayAlerts = false;
                    pivotWorkSheet.Delete();
                    pivotWorkSheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    pivotWorkSheet.Name = "Agent Reports";
                    pivotWorkSheet.Activate();
                }
                catch
                {
                    pivotWorkSheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    pivotWorkSheet.Name = "Agent Reports";
                    pivotWorkSheet.Activate();
                }

                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWindow.DisplayGridlines = false;

                #region Create Headings
                // Create headings
                // =================================================================================================================
                Excel.Range title = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "A1");
                title.ClearFormats();
                title.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color));
                title.Font.Color = System.Drawing.Color.White;
                title.InsertIndent(1);
                title.Font.Size = 18;
                title.VerticalAlignment = -4108; // xlCenter

                Excel.Range titleRow = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "L1");
                titleRow.Select();
                titleRow.RowHeight = 33;
                titleRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color));
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "D3").Font.Size = "10";
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "Agent Reports";
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 1] = "Generated by: " + userName;
                string s = mgmtServer; int start = s.IndexOf("//") + 2; int end = s.IndexOf(".", start); string serverName = s.Substring(start, end - start);
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 1] = "Server: " + serverName;

                Excel.Range rangeStamp = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 11];
                // rangeStamp.Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                rangeStamp.Value = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + " (UTC)";
                rangeStamp.NumberFormat = "[$-409]yyyy/mm/dd hh:mm AM/PM;@";
                rangeStamp.Cells.HorizontalAlignment = -4152; // Right align the Time
                rangeStamp.Value = rangeStamp.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text

                Excel.Range rowSeparator = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A3", "L3");
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color)); // 
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1; // xlContinuous
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = 4; // xlThick

                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "L3").Font.Size = "8";

                #endregion

                Globals.PivotBottom = 0;
                Globals.ChartBottom = 0;

                int NextPivotRowDefault = 0;
                int NextPivotRow = 0;
                int NextChartOffsetDefault = 0;
                int NextChartOffet = 0;

                Excel.Range rng = null;

                // Chart 1 =========================================================================================
                NextPivotRow = 5;
                NextChartOffet = 80;
                eHelper.CreatePivot("Agent Data", colCount, rowCount, "'Agent Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableOs", "Software Information.Os Name", "Operating System", "ID", "OS Count");
                eHelper.CreateChart("PivotTableOs", NextChartOffet, "Agent Operating Systems");


                // Chart 2 =========================================================================================
                NextPivotRowDefault = 20;
                NextChartOffsetDefault = 305;
                NextChartOffet = NextChartOffsetDefault;
                NextPivotRow = Globals.PivotBottom + 1;
                rng = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A" + NextPivotRow.ToString(), "A" + NextPivotRow.ToString());
                if (Globals.PivotBottom >= NextPivotRowDefault && Globals.ChartBottom < NextChartOffsetDefault)
                {
                    Globals.ChartBottom = (int)rng.Top;
                    NextChartOffet = Globals.ChartBottom;
                }
                else if (Globals.ChartBottom >= NextChartOffsetDefault)
                {
                    NextPivotRow = Globals.PivotBottom + 1;
                    bool doMore = true;
                    while (doMore)
                    {
                        rng = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A" + NextPivotRow.ToString(), "A" + NextPivotRow.ToString());
                        if (Globals.ChartBottom > rng.Top) NextPivotRow++;
                        else { doMore = false; }
                    }
                    NextChartOffet = (int)rng.Top;
                }
                else
                    NextPivotRow = NextPivotRowDefault;

                eHelper.CreatePivot("Agent Data", colCount, rowCount, "'Agent Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableAgent", "Agent Version", "Agent Version", "ID", "Version Count");
                eHelper.CreateChart("PivotTableAgent", NextChartOffet, "Agent Versions");

                // Chart 3 =========================================================================================
                NextPivotRowDefault = 35;
                NextChartOffsetDefault = 530;
                NextChartOffet = NextChartOffsetDefault;
                NextPivotRow = Globals.PivotBottom + 1;
                rng = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A" + NextPivotRow.ToString(), "A" + NextPivotRow.ToString());
                if (Globals.PivotBottom >= NextPivotRowDefault && Globals.ChartBottom < NextChartOffsetDefault)
                {
                    Globals.ChartBottom = (int)rng.Top;
                    NextChartOffet = Globals.ChartBottom;
                }
                else if (Globals.ChartBottom >= NextChartOffsetDefault)
                {
                    NextPivotRow = Globals.PivotBottom + 1;
                    bool doMore = true;
                    while (doMore)
                    {
                        rng = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A" + NextPivotRow.ToString(), "A" + NextPivotRow.ToString());
                        if (Globals.ChartBottom > rng.Top) NextPivotRow++;
                        else { doMore = false; }
                    }
                    NextChartOffet = (int)rng.Top;
                }
                else
                    NextPivotRow = NextPivotRowDefault;

                eHelper.CreatePivot("Agent Data", colCount, rowCount, "'Agent Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableIsActive", "Is Active", "Is Active", "ID", "Active Count");
                eHelper.CreateChart("PivotTableIsActive", NextChartOffet, "Agents Currently Active");

                // Chart 4 =========================================================================================
                NextPivotRowDefault = 50;
                NextChartOffsetDefault = 755;
                NextChartOffet = NextChartOffsetDefault;
                NextPivotRow = Globals.PivotBottom + 1;
                rng = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A" + NextPivotRow.ToString(), "A" + NextPivotRow.ToString());
                if (Globals.PivotBottom >= NextPivotRowDefault && Globals.ChartBottom < NextChartOffsetDefault)
                {
                    Globals.ChartBottom = (int)rng.Top;
                    NextChartOffet = Globals.ChartBottom;
                }
                else if (Globals.ChartBottom >= NextChartOffsetDefault)
                {
                    NextPivotRow = Globals.PivotBottom + 1;
                    bool doMore = true;
                    while (doMore)
                    {
                        rng = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A" + NextPivotRow.ToString(), "A" + NextPivotRow.ToString());
                        if (Globals.ChartBottom > rng.Top) NextPivotRow++;
                        else { doMore = false; }
                    }
                    NextChartOffet = (int)rng.Top;
                }
                else
                    NextPivotRow = NextPivotRowDefault;

                eHelper.CreatePivot("Agent Data", colCount, rowCount, "'Agent Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableGroupName", "Group Name", "Group Name", "ID", "Group Count");
                eHelper.CreateChart("PivotTableGroupName", NextChartOffet, "Groups with Most Agents");

            }
            catch (Exception ex)
            {
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = true;

                if (ex.Data.Contains("ExcelHelper"))
                {
                    MessageBox.Show(ex.Data["ExcelHelper"].ToString(), "Error generating chart", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(ex.Message, "Error generating chart", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                #region Finally
                if (pivotWorkSheet != null)
                    Marshal.ReleaseComObject(pivotWorkSheet);
                if (activeWorkBook != null)
                    Marshal.ReleaseComObject(activeWorkBook);
                #endregion
            }
        }

        #region Utilities
        private void AddFormatter(int col, string dataType, string key)
        {
            if (dataType == "DateTime")
            {
                dateTimeCol.Add(col);
            }
            else if (dataType == "Integer" && key != "mitigation_status")
            {
                integerCol.Add(col);
            }
            else if (dataType == "Reference")
            {
                referenceCol.Add(col);
            }
        }

        #endregion

        private void checkBoxAll_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAll.Checked)
            {
                textBoxQuery.Enabled = false;
            }
            else
            {
                textBoxQuery.Enabled = true;
            }

        }
    }
}
