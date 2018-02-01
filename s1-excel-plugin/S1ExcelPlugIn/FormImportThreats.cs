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
using System.Threading.Tasks;


namespace S1ExcelPlugIn
{
    public partial class FormImportThreats : Form
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

        int UnresolvedActive = 0;
        bool ActiveThreat = false;
        bool UnresolvedThreat = false;

        #endregion

        public FormImportThreats()
        {
            #region Initialization
            InitializeComponent();

            try
            {
                if (crypto.GetSettings("ExcelPlugIn", "ManagementServer") == null || crypto.GetSettings("ExcelPlugIn", "Token") == null)
                {
                    MessageBox.Show("Please log into the SentinelOne Management Server first", "Please login", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Load += (s, e) => Close();
                    return;
                }

                // Restoring previously specified Threat import parameters
                if (crypto.GetSettings("ExcelPlugIn", "MaximumRecord") == null)
                    numericUpDownMaxRecord.Value = 100;
                else
                    numericUpDownMaxRecord.Value = Convert.ToUInt32(crypto.GetSettings("ExcelPlugIn", "MaximumRecord"));

                if (crypto.GetSettings("ExcelPlugIn", "NoDateLimit") == null)
                    checkBoxDateLimit.Checked = false;
                else
                    checkBoxDateLimit.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "NoDateLimit"));

                if (crypto.GetSettings("ExcelPlugIn", "NoRecordLimit") == null)
                    checkBoxRecordLimit.Checked = false;
                else
                    checkBoxRecordLimit.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "NoRecordLimit"));

                if (crypto.GetSettings("ExcelPlugIn", "UnresolvedOnly") == null)
                    checkBoxResolved.Checked = false;
                else
                    checkBoxResolved.Checked = Convert.ToBoolean(crypto.GetSettings("ExcelPlugIn", "UnresolvedOnly"));

                if (crypto.GetSettings("ExcelPlugIn", "StartDate") == null)
                {
                    int year = DateTime.Now.Year;
                    DateTime firstDay = new DateTime(year, 1, 1);
                    dateTimePickerStart.Value = firstDay;
                }
                else
                {
                    dateTimePickerStart.Value = DateTime.Parse(crypto.GetSettings("ExcelPlugIn", "StartDate"));
                }

                if (crypto.GetSettings("ExcelPlugIn", "EndDate") == null)
                {
                    dateTimePickerEnd.Value = DateTime.Today;
                }
                else
                {
                    dateTimePickerEnd.Value = DateTime.Parse(crypto.GetSettings("ExcelPlugIn", "EndDate"));
                }

                double diff = ((dateTimePickerEnd.Value - dateTimePickerStart.Value).TotalDays + 1);
                textBoxDays.Text = diff == 1 ? diff.ToString() + " day" : diff.ToString() + " days";
                Globals.ReportPeriod = textBoxDays.Text;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error initialzing saved parameters, no harm done", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            #endregion
        }

        #region Button events
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
            if (GetData())
            {
                formMsg.Message("Creating pivot tables and charts for the threat data...", "", allowCancel: false);

                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = false;
                GenerateReport();
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = true;

                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = false;
                GenerateChart();
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = true;

                formMsg.Hide();

            }
            Close();
        }

        #endregion

        public bool GetData()
        {
            try
            {
                #region Create worksheet
                // Creates a worksheet for "Threats" if one does not already exist
                // =================================================================================================================
                Excel.Worksheet threatsWorksheet;
                try
                {
                    threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Threat Data");
                    threatsWorksheet.Activate();
                }
                catch
                {
                    threatsWorksheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    threatsWorksheet.Name = "Threat Data";
                    threatsWorksheet.Activate();
                }
                #endregion

                #region Clear worksheet
                // Clear spreadsheet
                eHelper.Clear();
                #endregion

                #region Get Data

                #region Get data from server
                // Extract data from SentinelOne Management Server
                // =================================================================================================================
                mgmtServer = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                string token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                userName = crypto.GetSettings("ExcelPlugIn", "Username");

                int MaxRecords = 1000000;
                if (checkBoxRecordLimit.Checked)
                {
                    // Leave MaxRecords with default
                }
                else
                {
                    MaxRecords = Convert.ToInt32(numericUpDownMaxRecord.Value);
                }
                int BatchRecords = Globals.ApiBatch;

                if (MaxRecords <= BatchRecords)
                {
                    BatchRecords = MaxRecords;
                }

                string created_at__lte = "";
                string created_at__gte = "";
                string daterange = "";

                crypto.SetSettings("ExcelPlugIn", "StartDate", dateTimePickerStart.Text);
                crypto.SetSettings("ExcelPlugIn", "EndDate", dateTimePickerEnd.Text);
                crypto.SetSettings("ExcelPlugIn", "NoDateLimit", checkBoxDateLimit.Checked.ToString());
                crypto.SetSettings("ExcelPlugIn", "MaximumRecord", numericUpDownMaxRecord.Text);
                crypto.SetSettings("ExcelPlugIn", "NoRecordLimit", checkBoxRecordLimit.Checked.ToString());
                crypto.SetSettings("ExcelPlugIn", "UnresolvedOnly", checkBoxResolved.Checked.ToString());

                Globals.StartDate = dateTimePickerStart.Text;
                Globals.EndDate = dateTimePickerEnd.Text;
                Globals.ReportPeriod = textBoxDays.Text;

                if (checkBoxDateLimit.Checked)
                {
                    daterange = "";
                    Globals.ReportPeriod = "All Time";
                    Globals.StartDate = "None";
                    Globals.EndDate = "None";
                }
                else
                {
                    created_at__gte = "created_at__gte=" + eHelper.DateTimeToUnixTimestamp(dateTimePickerStart.Value).ToString();
                    created_at__lte = "created_at__lte=" + eHelper.DateTimeToUnixTimestamp(dateTimePickerEnd.Value.AddDays(1)).ToString();
                    daterange = "&" + created_at__lte + "&" + created_at__gte;
                }

                string resolved = "";

                if (checkBoxResolved.Checked)
                {
                    resolved = "&resolved=false";
                    Globals.UnresolvedThreatOnly = true;
                }
                else
                {
                    Globals.UnresolvedThreatOnly = false;
                }

                string limit = "limit=" + BatchRecords.ToString();
                bool Gogo = true;
                string skip = "&skip=";
                int skip_count = 0;
                string results = "";
                int maxColumnWidth = 80;
                dynamic threats = "";
                int rowCountTemp = 0;
                string resourceString = "";

                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };

                var restClient = new RestClientInterface();

                stopWatch.Start();

                formMsg.StopProcessing = false;
                formMsg.Message("Loading threat data: " + rowCount.ToString("N0"),
                        eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);

                this.Opacity = 0;
                this.ShowInTaskbar = false;

                while (Gogo)
                {
                    resourceString = mgmtServer + "/web/api/v1.6/threats?" + limit + skip + skip_count.ToString() + daterange + resolved;
                    restClient.EndPoint = resourceString;
                    restClient.Method = HttpVerb.GET;
                    var batch_string = restClient.MakeRequest(token).ToString();
                    threats = Newtonsoft.Json.JsonConvert.DeserializeObject(batch_string, JsonSettings);
                    rowCountTemp = (int)threats.Count;
                    skip_count = skip_count + BatchRecords;
                    results = results + ", " + batch_string.TrimStart('[').TrimEnd(']', '\r', '\n');

                    rowCount = rowCount + rowCountTemp;

                    if (rowCountTemp < BatchRecords || rowCount >= MaxRecords)
                        Gogo = false;

                    if ((MaxRecords - rowCount) < BatchRecords)
                    {
                        limit = "limit=" + (MaxRecords - rowCount).ToString();
                    }

                    formMsg.Message("Loading threat data: " + rowCount.ToString("N0"),
                            eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);

                    if (formMsg.StopProcessing == true)
                    {
                        formMsg.Hide();
                        return false;
                    }

                }

                stopWatch.Stop();
                // formMsg.Hide();

                Globals.ApiUrl = resourceString;
                Globals.TotalThreats = rowCount;
                results = "[" + results.TrimStart(',').TrimEnd(',', ' ') + "]";
                threats = Newtonsoft.Json.JsonConvert.DeserializeObject(results, JsonSettings);

                #endregion

                #region Parse attribute headers
                JArray ja = JArray.Parse(results);
                // Stop processing if no data found
                if (ja.Count == 0)
                {
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "No threat data found";
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Select();
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Columns.AutoFit();
                    formMsg.Hide();
                    return false;
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
                                    AddFormatter(AttribNo - 1, dataType, Level1.Key + "." + Level2.Key + "." + Level3.Key);
                                }
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
                    else if (Level1.Key == "agent")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("agent_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "agent_name");

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("user_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "user_name");

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("group_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "group_name");

                    }
                    else if (Level1.Key == "classifier_name")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("detection_engine", "String"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "detection_engine");
                    }
                    else
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);
                    }
                }
                colCount = AttribCollection.Count;
                #endregion

                #endregion

                #region Create Headings
                // Create headings
                // =================================================================================================================
                Excel.Range titleRow = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", eHelper.ExcelColumnLetter(colCount - 1) + "1");
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
                string classifier = "";
                string fromCloud = "";
                string description = "";

                string[] prop;
                string header0 = "";
                string header1 = "";
                string header2 = "";
                string header3 = "";
                int headerLength = 0;

                Excel.Worksheet lookupSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                int rows = Convert.ToInt32(lookupSheet.Cells[3, 6].Value) + 4;

                formMsg.Message(rowCount.ToString("N0") + " threat data loaded, now processing locally in Excel...", "", allowCancel: false);

                for (int i = 0; i < rowCount; i++)
                {
                    ActiveThreat = false;
                    UnresolvedThreat = false;

                    for (int j = 0; j < AttribCollection.Count; j++)
                    {
                        prop = AttribCollection[j].Key.Split('.');
                        headerLength = prop.Length;
                        header0 = prop[0];
                        header1 = (headerLength > 1) ? prop[1] : "";
                        header2 = (headerLength > 2) ? prop[2] : "";
                        header3 = (headerLength > 3) ? prop[3] : "";

                        #region Nesting
                        if (prop.Length == 1)
                        {
                            temp = (threats[i][header0] == null) ? "null" : threats[i][header0].ToString();
                            if (header0 == "classifier_name")
                                classifier = temp;
                            else if (header0 == "from_cloud")
                                fromCloud = temp;
                            else if (header0 == "description")
                                description = temp;
                        }
                        else if (prop.Length == 2 && header0 == "engine_data")
                        {
                            if (threats[i][header0].ToString() == "[]")
                            {
                                temp = "null";
                                continue;
                            }
                            temp = (threats[i][header0][0] == null || threats[i][header0][0][header1] == null) ? "null" : threats[i][header0][0][header1].ToString();
                        }
                        else if (prop.Length == 2)
                            temp = (threats[i][header0][header1] == null) ? "null" : threats[i][header0][header1].ToString();
                        else if (prop.Length == 3)
                            temp = (threats[i][header0][header1][header2] == null) ? "null" : threats[i][header0][header1][header2].ToString();
                        else if (prop.Length == 4)
                            temp = (threats[i][header0][header1][header2][header3] == null) ? "null" : threats[i][header0][header1][header2][header3].ToString();
                        #endregion

                        #region Switch Statement

                        switch (header0)
                        {
                            #region mitigation_status
                            case "mitigation_status":
                                {
                                    switch ((string)temp)
                                    {
                                        case "0":
                                            temp = "Mitigated";
                                            break;
                                        case "1":
                                            temp = "Active";
                                            ActiveThreat = true;
                                            if (UnresolvedThreat)
                                            {
                                                UnresolvedActive++;
                                            }
                                            break;
                                        case "2":
                                            temp = "Blocked";
                                            break;
                                        case "3":
                                            temp = "Suspicious";
                                            break;
                                        case "4":
                                            temp = "Pending";
                                            break;
                                        case "5":
                                            temp = "Suspicious cancelled";
                                            break;
                                        default:
                                            temp = "Mitigated";
                                            break;
                                    }
                                    break;
                                }
                            #endregion

                            #region detection_engine
                            case "detection_engine":
                                {
                                    switch (classifier.ToLower())
                                    {
                                        case "blacklist":
                                            temp = "Pre-Execution (Blacklist)";
                                            break;
                                        case "preemptive":
                                            temp = "Pre-Execution (Blacklist)";
                                            break;
                                        case "static":
                                            if (fromCloud.ToLower() == "true")
                                                temp = "Pre-Execution (Cloud Reputation)";
                                            else
                                                temp = "Pre-Execution (DFI)";
                                            break;
                                        case "svm":
                                            temp = "On-Execution (DBT)";
                                            break;
                                        case "logic":
                                            temp = "On-Execution (DBT)";
                                            break;
                                        case "null":
                                            if (fromCloud.ToLower() == "true" || description.ToLower().Contains("sentinel cloud"))
                                                temp = "Pre-Execution (Cloud Reputation)";
                                            else
                                                temp = "Pre-Execution (DFI)";
                                            break;
                                        default:
                                            temp = "Unknown";
                                            break;
                                    }
                                    break;
                                }
                            #endregion

                            #region Lookups
                            case "agent_name":
                                {
                                    temp = "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(j - 1) + (i + 5).ToString() + ",'Lookup Tables'!F4:G" + rows.ToString() + ",2,FALSE),\"[Decommissioned]\")";
                                    break;
                                }

                            case "user_name":
                                {
                                    temp = "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(j - 2) + (i + 5).ToString() + ",'Lookup Tables'!F4:I" + rows.ToString() + ",4,FALSE),\"[Not Available]\")";
                                    break;
                                }

                            case "group_name":
                                {
                                    temp = "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(j - 3) + (i + 5).ToString() + ",'Lookup Tables'!F4:K" + rows.ToString() + ",6,FALSE),\"[Not Available]\")";
                                    break;
                                }
                            #endregion

                            #region resolved
                            case "resolved":
                                {
                                    if (temp.ToString().ToLower() == "false" && ActiveThreat)
                                    {
                                        UnresolvedActive++;
                                    }

                                    if (temp.ToString().ToLower() == "false")
                                    {
                                        UnresolvedThreat = true;
                                    }

                                    break;
                                }
                            #endregion

                            #region Date formatting
                            case "created_date":
                            case "meta_data":
                            case "last_active_date":
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

                        #endregion

                        dataBlock[i, j] = temp;
                    }
                } 

                Globals.ActiveAndUnresolvedThreats = UnresolvedActive.ToString();

                Excel.Range range;
                range = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A5", Missing.Value);

                if (threats.Count > 0)
                {
                    range = range.Resize[rowCount, colCount];
                    range.Cells.ClearFormats();
                    range.Value = dataBlock;
                    // range.Formula = range.Value;
                    range.Font.Size = "10";
                    range.RowHeight = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Height;
                }
                #endregion

                #region Column Headers
                // Column Headers
                // =================================================================================================================
                eHelper.WriteHeaders("Threat Data", colCount, rowCount, stopWatch);

                // This block writes the column headers
                string hd = "";
                for (int m = 0; m < AttribCollection.Count; m++)
                {
                    TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
                    hd = textInfo.ToTitleCase(AttribCollection[m].Key).Replace('_', ' ');
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[4, m + 1] = hd;
                }
                #endregion

                #region Worksheet formatting

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
                    // MessageBox.Show("Found DateTime");
                    Excel.Range rangeTime;
                    rangeTime = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range(eHelper.ExcelColumnLetter((int)dateTimeCol[j]) + "5", eHelper.ExcelColumnLetter((int)dateTimeCol[j]) + (rowCount + 5).ToString());
                    rangeTime.NumberFormat = "[$-409]yyyy/mm/dd hh:mm AM/PM;@";
                    // rangeTime.NumberFormat = "[$-800]yyyy/mm/dd hh:mm AM/PM;@";
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

                // Create the data filter
                // =================================================================================================================
                Excel.Range range2;
                range2 = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A4", Missing.Value);
                range2 = range2.Resize[rowCount + 1, colCount];
                range2.Select();
                // range2.AutoFilter("1", "<>", Excel.XlAutoFilterOperator.xlOr, "", true);
                range2.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlOr, Type.Missing, true);

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

                // This is saved for viewing in the Show API window
                Globals.ApiResults = JArray.Parse(results.ToString()).ToString();

                return true;
            }
            catch (Exception ex)
            {
                formMsg.Hide();
                MessageBox.Show(ex.Message, "Error extracting Threat data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void GenerateReport()
        {
            #region Initialize
            Excel.Workbook activeWorkBook = null;
            Excel.Worksheet pivotWorkSheet = null;
            Excel.PivotCaches pivotCaches = null;
            Excel.PivotCache pivotCache = null;
            Excel.PivotTable pivotTable = null;
            Excel.PivotFields pivotFields = null;

            Excel.PivotField monthPivotField = null;
            Excel.PivotField statusPivotField = null;
            Excel.PivotField resolvedPivotField = null;
            Excel.PivotField groupPivotField = null;
            Excel.PivotField threatIdPivotField = null;
            Excel.PivotField threatIdCountPivotField = null;

            Excel.SlicerCaches slicerCaches = null;

            Excel.SlicerCache monthSlicerCache = null;
            Excel.Slicers monthSlicers = null;
            Excel.Slicer monthSlicer = null;

            Excel.SlicerCache statusSlicerCache = null;
            Excel.Slicers statusSlicers = null;
            Excel.Slicer statusSlicer = null;

            Excel.SlicerCache resolvedSlicerCache = null;
            Excel.Slicers resolvedSlicers = null;
            Excel.Slicer resolvedSlicer = null;

            Excel.SlicerCache groupSlicerCache = null;
            Excel.Slicers groupSlicers = null;
            Excel.Slicer groupSlicer = null;
            #endregion

            try
            {
                activeWorkBook = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook;
                try
                {
                    pivotWorkSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Threat Reports");
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet.Application.DisplayAlerts = false;
                    pivotWorkSheet.Delete();
                    pivotWorkSheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    pivotWorkSheet.Name = "Threat Reports";
                    pivotWorkSheet.Activate();
                }
                catch
                {
                    pivotWorkSheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    pivotWorkSheet.Name = "Threat Reports";
                    pivotWorkSheet.Activate();
                }


                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWindow.DisplayGridlines = false;

                // Create the Pivot Table
                pivotCaches = activeWorkBook.PivotCaches();
                activeWorkBook.ShowPivotTableFieldList = false;
                string rangeName = "'Threat Data'!$A$4:$" + eHelper.ExcelColumnLetter(colCount-1) + "$" + (rowCount + 4).ToString();
                pivotCache = pivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, rangeName);
                pivotTable = pivotCache.CreatePivotTable("'Threat Reports'!R8C11", "PivotTableMitigationStatus");
                pivotTable.ClearTable();
                pivotTable.HasAutoFormat = false;
                pivotTable.NullString = "0";
                // Customizing the pivot table style
                pivotTable.TableStyle2 = "PivotStyleMedium9";

                // Set the Pivot Fields
                pivotFields = (Excel.PivotFields)pivotTable.PivotFields();
                // Month Pivot Field
                monthPivotField = (Excel.PivotField)pivotFields.Item("Created Date");
                monthPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                monthPivotField.Position = 1;
                monthPivotField.DataRange.Cells[1].Group(true, true, Type.Missing, new bool[] { false, false, false, false, true, true, true });
                pivotTable.CompactLayoutRowHeader = "Date Range";
                pivotTable.CompactLayoutColumnHeader = "Mitigation Status";
                // Mitigation Status Pivot Field
                statusPivotField = (Excel.PivotField)pivotFields.Item("Mitigation Status");
                statusPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                // Resolved Pivot Field
                resolvedPivotField = (Excel.PivotField)pivotFields.Item("Resolved");
                resolvedPivotField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                // Group Pivot Field
                groupPivotField = (Excel.PivotField)pivotFields.Item("Group Name");
                groupPivotField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
                // Threat ID Pivot Field
                threatIdPivotField = (Excel.PivotField)pivotFields.Item("ID");
                // Count of Threat ID Field
                threatIdCountPivotField = pivotTable.AddDataField(threatIdPivotField, "# of Threats", Excel.XlConsolidationFunction.xlCount);
                slicerCaches = activeWorkBook.SlicerCaches;
                // Month Slicer
                monthSlicerCache = slicerCaches.Add(pivotTable, "Created Date", "CreatedDate");
                monthSlicers = monthSlicerCache.Slicers;
                // monthSlicer = monthSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Created Date", "Created Date", 60, 10, 144, 100);
                monthSlicer = monthSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Created Date", "Created Date", 60, 10, 85, 100);
                // Mitigation Status Slicer
                statusSlicerCache = slicerCaches.Add(pivotTable, "Mitigation Status", "MitigationStatus");
                statusSlicers = statusSlicerCache.Slicers;
                // statusSlicer = statusSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Mitigation Status", "Mitigation Status", 60, 164, 144, 100);
                statusSlicer = statusSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Mitigation Status", "Mitigation Status", 60, 105, 105, 100);
                // Resolved Slicer
                resolvedSlicerCache = slicerCaches.Add(pivotTable, "Resolved", "Resolved");
                resolvedSlicers = resolvedSlicerCache.Slicers;
                // resolvedSlicer = resolvedSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Resolved", "Resolved", 60, 318, 144, 100);
                resolvedSlicer = resolvedSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Resolved", "Resolved", 60, 220, 85, 100);
                // Group Slicer
                groupSlicerCache = slicerCaches.Add(pivotTable, "Group Name", "GroupName");
                groupSlicers = groupSlicerCache.Slicers;
                groupSlicer = groupSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Group Name", "Group Name", 60, 315, 145, 100);

                int iTotalColumns = pivotWorkSheet.UsedRange.Columns.Count + 9;
                int iTotalRows = pivotWorkSheet.UsedRange.Rows.Count;

                // Remembers the bottom of the pivot table so that the next one will not overlap
                Globals.PivotBottom = 0;
                Globals.ChartBottom = 0;
                Globals.PivotBottom = pivotTable.TableRange2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Row + pivotTable.TableRange2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Count;

                #region Create Headings
                // Create headings
                // ================================================================================================================
                Excel.Range title = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "A1");
                title.ClearFormats();
                title.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color));
                title.Font.Color = System.Drawing.Color.White;
                title.InsertIndent(1);
                title.Font.Size = 18;
                title.VerticalAlignment = -4108; // xlCenter

                Excel.Range titleRow = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", eHelper.ExcelColumnLetter(iTotalColumns) + "1");
                titleRow.Select();
                titleRow.RowHeight = 33;
                titleRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color));
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "D3").Font.Size = "10";
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "Threat Reports";
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 1] = "Generated by: " + userName;
                string s = mgmtServer; int start = s.IndexOf("//") + 2; int end = s.IndexOf(".", start); string serverName = s.Substring(start, end - start);
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 1] = "Server: " + serverName;

                Excel.Range rangeStamp = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 11];
                rangeStamp.Value = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") + " (UTC)";
                rangeStamp.NumberFormat = "[$-409]yyyy/mm/dd hh:mm AM/PM;@";
                rangeStamp.Cells.HorizontalAlignment = -4152; // Right align the Time
                rangeStamp.Value = rangeStamp.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text

                Excel.Range rowSeparator = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A3", eHelper.ExcelColumnLetter(iTotalColumns) + "3");
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color)); // 
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1; // xlContinuous
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = 4; // xlThick

                Excel.Range pivotRange = pivotTable.TableRange1;
                // MessageBox.Show("Col count: " + pivotRange.Columns.Count.ToString() + ", Row count: " + pivotRange.Rows.Count.ToString());

                string reportTable = "";
                string reportLabel = "\"";
                string reportValue = "";

                reportTable = "<table id=\"newspaper-a\" class=\"sortable\">";

                string head1 = "";
                string headString = "<thead><tr>";
                string body1 = "";
                string bodyString = "<tbody><tr>";

                for (int j = 1; j < pivotRange.Columns.Count; j++)
                {
                    head1 = (string)(pivotWorkSheet.Cells[pivotRange.Row + 1, pivotRange.Column + j] as Excel.Range).Value.ToString();
                    headString = headString + "<th scope=\"col\">" + head1 + "</th>";

                    body1 = (string)(pivotWorkSheet.Cells[pivotRange.Row + pivotRange.Rows.Count - 1, pivotRange.Column + j] as Excel.Range).Value.ToString();
                    bodyString = bodyString + "<td>" + body1 + "</td>";

                    if (head1.ToLower() == "active")
                    {
                        Globals.ActiveThreats = body1;
                        if (Globals.ActiveThreats == "")
                            Globals.ActiveThreats = "0";
                    }

                    reportLabel = reportLabel + head1 + "\",\"";
                    reportValue = reportValue + body1 + ",";
                }
                headString = headString + "</tr></thead>";
                bodyString = bodyString + "</tr></tbody>";

                reportValue = reportValue.TrimEnd(',');
                reportLabel = reportLabel.TrimEnd('\"');
                reportLabel = reportLabel.TrimEnd(',');

                reportTable = reportTable + headString + bodyString + "</table>";

                Globals.ThreatData = reportTable;
                Globals.ThreatDataLabel = reportLabel;
                Globals.ThreatDataValue = reportValue;

                // MessageBox.Show(reportTable + "\r\n\r\n" + reportLabel + "\r\n\r\n" + reportValue);

                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "L3").Font.Size = "8";
                #endregion

            }
            catch (Exception ex)
            {
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = true;
                MessageBox.Show(ex.Message, "Error generating report", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                #region Finally
                if (resolvedSlicer != null)
                    Marshal.ReleaseComObject(resolvedSlicer);
                if (resolvedSlicers != null)
                    Marshal.ReleaseComObject(resolvedSlicers);
                if (resolvedSlicerCache != null)
                    Marshal.ReleaseComObject(resolvedSlicerCache);
                if (statusSlicer != null)
                    Marshal.ReleaseComObject(statusSlicer);
                if (statusSlicers != null)
                    Marshal.ReleaseComObject(statusSlicers);
                if (statusSlicerCache != null)
                    Marshal.ReleaseComObject(statusSlicerCache);
                if (monthSlicer != null)
                    Marshal.ReleaseComObject(monthSlicer);
                if (monthSlicers != null)
                    Marshal.ReleaseComObject(monthSlicers);
                if (monthSlicerCache != null)
                    Marshal.ReleaseComObject(monthSlicerCache);
                if (slicerCaches != null)
                    Marshal.ReleaseComObject(slicerCaches);
                if (threatIdCountPivotField != null)
                    Marshal.ReleaseComObject(threatIdCountPivotField);
                if (threatIdPivotField != null)
                    Marshal.ReleaseComObject(threatIdPivotField);
                if (resolvedPivotField != null)
                    Marshal.ReleaseComObject(resolvedPivotField);
                if (statusPivotField != null)
                    Marshal.ReleaseComObject(statusPivotField);
                if (monthPivotField != null)
                    Marshal.ReleaseComObject(monthPivotField);
                if (pivotFields != null)
                    Marshal.ReleaseComObject(pivotFields);
                if (pivotTable != null)
                    Marshal.ReleaseComObject(pivotTable);
                if (pivotCache != null)
                    Marshal.ReleaseComObject(pivotCache);
                if (pivotCaches != null)
                    Marshal.ReleaseComObject(pivotCaches);
                if (pivotWorkSheet != null)
                    Marshal.ReleaseComObject(pivotWorkSheet);
                if (activeWorkBook != null)
                    Marshal.ReleaseComObject(activeWorkBook);
                #endregion
            }
        }

        private void GenerateChart()
        {
            Excel.Worksheet activeSheet = null;
            Excel.Range selectedRange = null;
            Excel.Shapes shapes = null;
            Excel.Chart chart = null;
            Excel.ChartTitle chartTitle = null;

            try
            {
                activeSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet;
                Excel.Range rangePivot;
                rangePivot = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "AA2000");
                selectedRange = eHelper.IdentifyPivotRangesByName("PivotTableMitigationStatus");
                shapes = activeSheet.Shapes;

                if (Globals.ExcelVersion == "15.0" || Globals.ExcelVersion == "16.0")
                {
                    shapes.AddChart2(Style: 340, XlChartType: Excel.XlChartType.xlColumnClustered,
                    Left: 10, Top: 190, Width: 450,
                    Height: Type.Missing, NewLayout: true).Select();
                }
                else
                {
                    shapes.AddChart(XlChartType: Excel.XlChartType.xlColumnClustered,
                    Left: 10, Top: 190, Width: 450,
                    Height: Type.Missing).Select();
                }

                chart = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveChart;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Detected Threat Statuses";
                chart.ChartArea.Interior.Color = System.Drawing.Color.FromArgb(242, 244, 244); // Change chart to light gray
                chart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowValue, true, true, true, true, true, true, true, true, true); // Turn on data labels
                chart.HasLegend = true;
                chart.SetSourceData(selectedRange);
                if (selectedRange.Rows.Count > 12)
                    chart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelNone);
                else
                    chart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelOutSideEnd);

                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeTop].ColorIndex = 0;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeTop].TintAndShade = 0;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlMedium;

                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeRight].ColorIndex = 0;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeRight].TintAndShade = 0;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].ColorIndex = 0;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].TintAndShade = 0;
                selectedRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

                // chart.Export("mitigation_status_chart.png");

                Globals.ChartBottom = (int)chart.ChartArea.Top + (int)chart.ChartArea.Height;
                Excel.Range rng = null;
                int NextPivotRowDefault = 0;
                int NextChartOffsetDefault = 0;
                int NextPivotRow = 0;
                int NextChartOffet = 0;

                #region Detection Engine
                // Threat classifier ====================================================================================================================
                NextPivotRowDefault = 28;
                NextChartOffsetDefault = 425;
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

                eHelper.CreatePivot("Threat Data", colCount, rowCount, "'Threat Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableClassifier", "Detection Engine", "Detection Engine", "Resolved", "Detection Count");
                eHelper.CreateChart("PivotTableClassifier", NextChartOffet, "SentinelOne Detection Engines");
                #endregion

                #region Detected File
                // Most detected file ====================================================================================================================
                NextPivotRowDefault = NextPivotRowDefault + 15;
                NextChartOffsetDefault = NextChartOffsetDefault + 225;
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
                {
                    NextPivotRow = NextPivotRowDefault;
                }

                eHelper.CreatePivot("Threat Data", colCount, rowCount, "'Threat Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableFileDisplayName", "File Id.Display Name", "File Id.Display Name", "Resolved", "Incident Count");
                eHelper.CreateChart("PivotTableFileDisplayName", NextChartOffet, "Most Convicted Files");
                #endregion

                
                #region At Risk Users
                // Most at-risk users ====================================================================================================================
                NextPivotRowDefault = NextPivotRowDefault + 15;
                NextChartOffsetDefault = NextChartOffsetDefault + 225;
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

                eHelper.CreatePivot("Threat Data", colCount, rowCount, "'Threat Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableAtRiskUsers", "User Name", "User Name", "Resolved", "Incident Count");
                eHelper.CreateChart("PivotTableAtRiskUsers", NextChartOffet, "Most At-Risk Users");
                #endregion


                #region At Risk Endpoints
                // Most at-risk endpoints ====================================================================================================================
                NextPivotRowDefault = NextPivotRowDefault + 15;
                NextChartOffsetDefault = NextChartOffsetDefault + 225;
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

                eHelper.CreatePivot("Threat Data", colCount, rowCount, "'Threat Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableAtRiskEndpoints", "Agent Name", "Agent Name", "Resolved", "Incident Count");
                eHelper.CreateChart("PivotTableAtRiskEndpoints", NextChartOffet, "Most At-Risk Endpoints");
                #endregion

                #region At Risk Groups
                // Most at-risk groups ====================================================================================================================
                NextPivotRowDefault = NextPivotRowDefault + 15;
                NextChartOffsetDefault = NextChartOffsetDefault + 225;
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

                eHelper.CreatePivot("Threat Data", colCount, rowCount, "'Threat Reports'!R" + NextPivotRow.ToString() + "C11", 
                                    "PivotTableAtRiskGroups", "Group Name", "Group Name", "Resolved", "Incident Count");
                eHelper.CreateChart("PivotTableAtRiskGroups", NextChartOffet, "Most At-Risk Groups");
                #endregion


                // Cursor placement ===================================================================
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Select();
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
                if (chartTitle != null) Marshal.ReleaseComObject(chartTitle);
                if (chart != null) Marshal.ReleaseComObject(chart);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
            }
        }

        #region Evant handlers
        private void dateTimePickerStart_ValueChanged(object sender, EventArgs e)
        {
            double diff = ((dateTimePickerEnd.Value - dateTimePickerStart.Value).TotalDays + 1);
            textBoxDays.Text = diff == 1 ? diff.ToString() + " day" : diff.ToString() + " days";
        }

        private void dateTimePickerEnd_ValueChanged(object sender, EventArgs e)
        {
            double diff = ((dateTimePickerEnd.Value - dateTimePickerStart.Value).TotalDays + 1);
            textBoxDays.Text = diff == 1 ? diff.ToString() + " day" : diff.ToString() + " days";
        }

        private void checkBoxDateLimit_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBoxDateLimit.Checked)
            {
                dateTimePickerStart.Enabled = false;
                dateTimePickerEnd.Enabled = false;
                textBoxDays.Enabled = false;
            }
            else
            {
                dateTimePickerStart.Enabled = true;
                dateTimePickerEnd.Enabled = true;
                textBoxDays.Enabled = true;
            }
        }

        private void checkBoxRecordLimit_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxRecordLimit.Checked)
            {
                numericUpDownMaxRecord.Enabled = false;
            }
            else
            {
                numericUpDownMaxRecord.Enabled = true;
            }
        }

        #endregion

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

    }
}
