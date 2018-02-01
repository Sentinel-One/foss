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

namespace S1ExcelPlugIn
{
    public partial class FormImportActivities : Form
    {
        #region Class variables
        Crypto crypto = new Crypto();

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

        #endregion

        public FormImportActivities()
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

                // Restoring previously specified Activites parameters
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

                // =============================================================================================
                string token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                string server = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                string server2 = server.Substring(server.LastIndexOf('/') + 1);
                string username = crypto.GetSettings("ExcelPlugIn", "Username");
                string resourceString = server + "/web/api/v1.6/activities/types";

                var restClient = new RestClientInterface(resourceString);
                restClient.Method = HttpVerb.GET;
                restClient.PostData = "";
                var results = restClient.MakeRequest(token).ToString();

                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };

                dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(results, JsonSettings);
                rowCount = (int)x.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    ActivityType at = new ActivityType();
                    at.Text = x[i].action.ToString().Replace("Ad ", "AD ");
                    at.Value = x[i].id.ToString();
                    listBoxTotal.Items.Add(at);
                }

                if (crypto.GetSettings("ExcelPlugIn", "ActivityTypes") != null)
                {
                    string act_types = crypto.GetSettings("ExcelPlugIn", "ActivityTypes");
                    string[] type_array = act_types.Split(',');
                    for (int i = 0; i < type_array.Length; i++)
                    {
                        for (int j = 0; j < listBoxTotal.Items.Count; j++)
                        {
                            if (listBoxTotal.Items[j].ToString() == type_array[i])
                            {
                                listBoxSelected.Items.Add(listBoxTotal.Items[j]);
                                listBoxTotal.Items.Remove(listBoxTotal.Items[j]);
                            }
                        }
                    }
                }

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
            /*
            GetData();
            GenerateReport();
            GenerateChart();
            Close();
            */
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            if (listBoxTotal.SelectedItems.Count > 0)
            {
                for (int i = 0; i < listBoxTotal.SelectedItems.Count; i++)
                {
                    listBoxSelected.Items.Add(listBoxTotal.SelectedItems[i]);
                }

                if (listBoxTotal.SelectedIndex != -1)
                {
                    for (int i = listBoxTotal.SelectedItems.Count - 1; i >= 0; i--)
                        listBoxTotal.Items.Remove(listBoxTotal.SelectedItems[i]);
                }
            }
            else
            {
                MessageBox.Show("Nothing selected", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (listBoxSelected.Items.Count == 0)
                listBoxTotal.Enabled = false;

        }

        private void buttonDelete_Click(object sender, EventArgs e)
        {
            if (listBoxSelected.SelectedItems.Count > 0)
            {
                for (int i = 0; i < listBoxSelected.SelectedItems.Count; i++)
                {
                    listBoxTotal.Items.Add(listBoxSelected.SelectedItems[i]);
                }

                if (listBoxSelected.SelectedIndex != -1)
                {
                    for (int i = listBoxSelected.SelectedItems.Count - 1; i >= 0; i--)
                        listBoxSelected.Items.Remove(listBoxSelected.SelectedItems[i]);
                }
            }
            else
            {
                MessageBox.Show("Nothing selected", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (listBoxSelected.Items.Count == 0)
                buttonUnselectAll.Enabled = false;

        }

        private void buttonSelectAll_Click(object sender, EventArgs e)
        {
            if (listBoxTotal.Items.Count > 0)
            {
                for (int i = 0; i < listBoxTotal.Items.Count; i++)
                {
                    listBoxSelected.Items.Add(listBoxTotal.Items[i]);
                }

                listBoxTotal.Items.Clear();
                buttonSelectAll.Enabled = false;
                buttonUnselectAll.Enabled = true;
            }
            else
            {
                MessageBox.Show("Nothing to move", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void buttonUnselectAll_Click(object sender, EventArgs e)
        {
            if (listBoxSelected.Items.Count > 0)
            {
                for (int i = 0; i < listBoxSelected.Items.Count; i++)
                {
                    listBoxTotal.Items.Add(listBoxSelected.Items[i]);
                }

                listBoxSelected.Items.Clear();
                buttonUnselectAll.Enabled = false;
                buttonSelectAll.Enabled = true;
            }
            else
            {
                MessageBox.Show("Nothing to move", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
                    threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Activity Data");
                    threatsWorksheet.Activate();
                }
                catch
                {
                    threatsWorksheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    threatsWorksheet.Name = "Activity Data";
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
                string mgmtServer = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                string token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));

                string created_at__lte = "";
                string created_at__gte = "";
                string daterange = "";
                string activity_type = "&activity_type__in=";

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

                int numberOfSelected = listBoxSelected.Items.Count;
                if (numberOfSelected > 0)
                {
                    for (int i = 0; i < numberOfSelected; i++)
                    {
                        activity_type = activity_type + listBoxSelected.Items[i].ToString() + ",";
                    }
                    activity_type = activity_type.TrimEnd(',');
                }
                else
                {
                    MessageBox.Show("No activity types selected", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (checkBoxDateLimit.Checked)
                {
                    daterange = "";
                }
                else
                {
                    created_at__gte = "created_at__gte=" + eHelper.DateTimeToUnixTimestamp(dateTimePickerStart.Value).ToString();
                    created_at__lte = "created_at__lte=" + eHelper.DateTimeToUnixTimestamp(dateTimePickerEnd.Value.AddDays(1)).ToString();
                    daterange = "&" + created_at__lte + "&" + created_at__gte;
                }

                crypto.SetSettings("ExcelPlugIn", "StartDate", dateTimePickerStart.Text);
                crypto.SetSettings("ExcelPlugIn", "EndDate", dateTimePickerEnd.Text);
                crypto.SetSettings("ExcelPlugIn", "NoDateLimit", checkBoxDateLimit.Checked.ToString());
                crypto.SetSettings("ExcelPlugIn", "ActivityTypes", activity_type.Substring(activity_type.IndexOf('=')+1));
                crypto.SetSettings("ExcelPlugIn", "MaximumRecord", numericUpDownMaxRecord.Text);
                crypto.SetSettings("ExcelPlugIn", "NoRecordLimit", checkBoxRecordLimit.Checked.ToString());

                string limit = "limit=" + BatchRecords.ToString();
                bool Gogo = true;
                string skip = "&skip=";
                int skip_count = 0;
                string results = "";
                int maxColumnWidth = 80;
                dynamic threats = "";
                rowCount = 0;
                int rowCountTemp = 0;
                string resourceString = "";

                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };

                var restClient = new RestClientInterface();

                formMsg.StopProcessing = false;
                formMsg.StartMessage();

                this.Opacity = 0;
                this.ShowInTaskbar = false;

                stopWatch.Start();

                while (Gogo)
                {
                    resourceString = mgmtServer + "/web/api/v1.6/activities?" + limit + skip + skip_count.ToString() + daterange + activity_type;
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

                    formMsg.UpdateMessage("Loading activity data: " + rowCount.ToString("N0"),
                            eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed");

                    if (formMsg.StopProcessing == true)
                    {
                        formMsg.Hide();
                        return;
                    }
                }

                stopWatch.Stop();
                formMsg.Hide();

                Globals.ApiUrl = resourceString;
                results = "[" + results.TrimStart(',').TrimEnd(',', ' ') + "]";
                threats = Newtonsoft.Json.JsonConvert.DeserializeObject(results, JsonSettings);
                #endregion

                #region Parse attribute headers
                JArray ja = JArray.Parse(results);
                // Stop processing if no data found
                if (ja.Count == 0)
                {
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "No activity data found";
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Select();
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Columns.AutoFit();
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
                    else if (Level1.Key == "activity_type")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("activity_type_description", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "activity_type_description");
                    }
                    else if (Level1.Key == "user_id")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("user_full_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "user_full_name");
                    }
                    else if (Level1.Key == "agent_id")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("agent_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "agent_name");
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
                Excel.Range titleRow = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", eHelper.ExcelColumnLetter(colCount-1) + "1");
                titleRow.Select();
                titleRow.RowHeight = 33;
                titleRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color));

                Excel.Range rowSeparator = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A3", eHelper.ExcelColumnLetter(colCount-1) + "3");
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color)); // 
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1; // xlContinuous
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = 4; // xlThick
                #endregion

                #region Write data
                // Write all the data rows
                // =================================================================================================================
                string[,] dataBlock = new string[rowCount, colCount];
                dynamic temp = "";
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < AttribCollection.Count; j++)
                    {
                        string[] prop = AttribCollection[j].Key.Split('.');

                        if (prop.Length == 1)
                            temp = (threats[i][prop[0]] == null) ? "null" : threats[i][prop[0]].ToString();
                        else if (prop.Length == 2 && prop[0] == "engine_data")
                        {
                            if (threats[i][prop[0]].ToString() == "[]")
                            {
                                temp = "null";
                                continue;
                            }
                            temp = (threats[i][prop[0]][0] == null || threats[i][prop[0]][0][prop[1]] == null) ? "null" : threats[i][prop[0]][0][prop[1]].ToString();
                        }
                        else if (prop.Length == 2)
                            temp = (threats[i][prop[0]][prop[1]] == null) ? "null" : threats[i][prop[0]][prop[1]].ToString();
                        else if (prop.Length == 3)
                            temp = (threats[i][prop[0]][prop[1]][prop[2]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]].ToString();
                        else if (prop.Length == 4)
                            temp = (threats[i][prop[0]][prop[1]][prop[2]][prop[3]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]][prop[3]].ToString();

                        if (prop[0] == "created_date" || prop[0] == "meta_data" || prop[0] == "last_active_date")
                        {
                            DateTime d2 = DateTime.Parse(temp, null, System.Globalization.DateTimeStyles.RoundtripKind);
                            temp = "=DATE(" + d2.ToString("yyyy,MM,dd") + ")+TIME(" + d2.ToString("HH,mm,ss") + ")";
                        }

                        try
                        {
                            if (prop[0] == "activity_type_description")
                            {
                                Excel.Worksheet lookupSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                                int rows = Convert.ToInt32(lookupSheet.Cells[3, 1].Value) + 4;
                                temp = "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(j - 1) + (i + 5).ToString() + ",'Lookup Tables'!A4:B" + rows.ToString() + ",2,FALSE),\"null\")";
                                // =VLOOKUP(N5,'Lookup Tables'!A4:B82,2,FALSE)
                            }

                            if (prop[0] == "agent_name")
                            {
                                Excel.Worksheet lookupSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                                int rows = Convert.ToInt32(lookupSheet.Cells[3, 6].Value) + 4;
                                temp = "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(j - 1) + (i + 5).ToString() + ",'Lookup Tables'!F4:G" + rows.ToString() + ",2,FALSE),\"null\")";
                                // =VLOOKUP(N5,'Lookup Tables'!A4:B82,2,FALSE)
                            }

                            if (prop[0] == "user_full_name")
                            {
                                Excel.Worksheet lookupSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                                int rows = Convert.ToInt32(lookupSheet.Cells[3, 15].Value) + 4;
                                temp = "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(j - 1) + (i + 5).ToString() + ",'Lookup Tables'!O4:P" + rows.ToString() + ",2,FALSE),\"null\")";
                            }
                        }
                        catch { }
                        dataBlock[i, j] = temp;
                    }
                }

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
                eHelper.WriteHeaders("Activity Data", colCount, rowCount, stopWatch);

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

                Globals.ApiResults = JArray.Parse(results.ToString()).ToString();

            }
            catch (Exception ex)
            {
                formMsg.Hide();
                MessageBox.Show(ex.Message, "Error extracting Activity data", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Event Handlers
        private void listBoxTotal_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxTotal.SelectedItems.Count > 0)
            {
                buttonAdd.Enabled = true;
            }
            else
            {
                buttonAdd.Enabled = false;
            }

            if (listBoxTotal.Items.Count > 0)
                buttonSelectAll.Enabled = true;
            else
                buttonSelectAll.Enabled = false;
        }

        private void listBoxSelected_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxSelected.SelectedItems.Count > 0)
            {
                buttonDelete.Enabled = true;
            }
            else
            {
                buttonDelete.Enabled = false;
            }

            if (listBoxSelected.Items.Count > 0)
                buttonUnselectAll.Enabled = true;
            else
                buttonUnselectAll.Enabled = false;

        }

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

    public class ActivityType
    {
        public string Value { get; set; }
        public string Text { get; set; }
        public override string ToString()
        {
            return Value;
        }
    }
}
