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
    public partial class FormImportProcesses : Form
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

        ExcelHelper eHelper = new ExcelHelper();

        System.Array AgentArray;
        int AgentArrayItems = 0;
        int UpdateInterval = Globals.ProcessBatch;
        int maxIterations = 10000;
        // int maxIterations = 100;

        FormMessages formMsg = new FormMessages();

        #endregion

        public FormImportProcesses()
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

                string procBatch = crypto.GetSettings("ExcelPlugIn", "ProcessBatch");

                if (procBatch == null)
                {
                    crypto.SetSettings("ExcelPlugIn", "ProcessBatch", UpdateInterval.ToString());
                }
                else
                {
                    UpdateInterval = Convert.ToInt32(procBatch);
                }
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

            formMsg.Message("Creating pivot tables and charts for the process data...", "", allowCancel: false);

            (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = false;
            GenerateReport();
            (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ScreenUpdating = true;

            formMsg.Hide();

            Close();
        }

        #endregion 

        public void GetData()
        {
            try
            {
                #region Create worksheet
                // Creates a worksheet for "Applications" if one does not already exist
                // =================================================================================================================
                Excel.Worksheet threatsWorksheet;
                try
                {
                    threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Process Data");
                    threatsWorksheet.Activate();
                }
                catch
                {
                    threatsWorksheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    threatsWorksheet.Name = "Process Data";
                    threatsWorksheet.Activate();
                }
                #endregion

                #region Clear worksheet
                // Clear spreadsheet
                eHelper.Clear();
                #endregion

                #region Get Data

                #region Find All the Agents from the hidden Lookup Table sheet
                Excel.Worksheet lookupSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                int rows = Convert.ToInt32(lookupSheet.Cells[3, 6].Value) + 4;
                Excel.Range AgentIDs = lookupSheet.Range["F5:H" + rows.ToString()];
                AgentArray = AgentIDs.Value;
                AgentArrayItems = AgentArray.GetLength(0);
                #endregion

                #region Get data from server (Looping Call)

                // Extract data from SentinelOne Management Server (Looping Calls)
                // =================================================================================================================
                mgmtServer = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                string token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                userName = crypto.GetSettings("ExcelPlugIn", "Username");

                StringBuilder results = new StringBuilder("[");
                int maxColumnWidth = 80;
                dynamic threats = "";
                int rowCountTemp = 0;
                int percentComplete = 0;
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };

                string resourceString = "";
                var restClient = new RestClientInterface();

                if (AgentArrayItems > maxIterations)
                    AgentArrayItems = maxIterations;

                formMsg.StopProcessing = false;
                formMsg.Message("Loading process data: " + rowCount.ToString("N0"),
                        eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);

                for (int i = 1; i <= AgentArrayItems; i++)
                {
                    resourceString = mgmtServer + "/web/api/v1.6/agents/" + AgentArray.GetValue(i,1).ToString() + "/processes";

                    Globals.ApiUrl = resourceString;
                    restClient.EndPoint = resourceString;
                    restClient.Method = HttpVerb.GET;

                    var batch_string = "";
                    try
                    {
                        batch_string = restClient.MakeRequest(token).ToString();
                    }
                    catch { };

                    if (batch_string.Length < 4)
                    {
                        continue;
                    }

                    if (i % UpdateInterval == 0)
                    {
                        percentComplete = (int)Math.Round((double)(100 * i) / AgentArrayItems);
                        formMsg.Message("Collecting processes from " + i.ToString("N0") + " of " + AgentArrayItems.ToString("N0") + " agents (" + percentComplete.ToString() + "%)...",
                            rowCount.ToString("N0") + " processes found, " + eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);
                        if (formMsg.StopProcessing == true)
                        {
                            formMsg.Hide();
                            return;
                        }
                    }

                    JArray batchArray = JArray.Parse(batch_string);
                    string procName = "";

                    List<string> excludeProcesses = new List<string>();
                    excludeProcesses.Add("conhost.exe");
                    excludeProcesses.Add("csrss.exe");
                    excludeProcesses.Add("dllhost.exe");
                    excludeProcesses.Add("dwm.exe");
                    excludeProcesses.Add("explorer.exe");
                    excludeProcesses.Add("logonui.exe");
                    excludeProcesses.Add("lsass.exe");
                    excludeProcesses.Add("msdtc.exe");
                    excludeProcesses.Add("rundll32.exe");
                    excludeProcesses.Add("runtimebroker.exe");
                    excludeProcesses.Add("searchindexer.exe");
                    excludeProcesses.Add("searchprotocolhost.exe");
                    excludeProcesses.Add("sentinelagent.exe");
                    excludeProcesses.Add("sentinelagent.exe");
                    excludeProcesses.Add("sentinelagent.exe");
                    excludeProcesses.Add("sentinelservicehost.exe");
                    excludeProcesses.Add("sentinelstaticengine.exe");
                    excludeProcesses.Add("sentinelstaticenginescanner.exe");
                    excludeProcesses.Add("services.exe");
                    excludeProcesses.Add("smss.exe");
                    excludeProcesses.Add("smsvchost.exe");
                    excludeProcesses.Add("spoolsv.exe");
                    excludeProcesses.Add("svchost.exe");
                    excludeProcesses.Add("taskhost.exe");
                    excludeProcesses.Add("taskhostw.exe");
                    excludeProcesses.Add("wininit.exe");
                    excludeProcesses.Add("winlogon.exe");
                    excludeProcesses.Add("wmiprvse.exe");
                    excludeProcesses.Add("wudfhost.exe");

                    for (int j = batchArray.Count; j --> 0;)
                    {
                        procName = ((JObject)batchArray[j])["process_name"].ToString();

                        if (excludeProcesses.Contains(procName.ToLower()))
                        {
                            batchArray[j].Remove();
                        }
                        else
                        {
                            var PropertyComputerName = new JProperty("computer_name", AgentArray.GetValue(i, 2) == null ? "N/A" : AgentArray.GetValue(i, 2).ToString());
                            var PropertyOperatingSystem = new JProperty("operating_system", AgentArray.GetValue(i, 3) == null ? "N/A" : AgentArray.GetValue(i, 3).ToString());

                            ((JObject)batchArray[j]).AddFirst(PropertyOperatingSystem);
                            ((JObject)batchArray[j]).AddFirst(PropertyComputerName);
                        }

                    }

                    batch_string = batchArray.ToString();

                    threats = Newtonsoft.Json.JsonConvert.DeserializeObject(batch_string, JsonSettings);
                    rowCountTemp = (int)threats.Count;
                    results.Append(batchArray.ToString().TrimStart('[').TrimEnd(']', '\r', '\n')).Append(",");
                    rowCount = rowCount + rowCountTemp;
                }
                stopWatch.Stop();
                formMsg.Hide();

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
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "No process data found";
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
                            else
                            {
                                AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + Level2.Key, dataType));
                                AttribNo++;
                                AddFormatter(AttribNo - 1, dataType, Level1.Key + "." + Level2.Key);
                            }
                        }
                    }
                    else if (Level1.Key == "computer_name")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("computer_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "computer_name");
                    }
                    else if (Level1.Key == "operating_system")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("operating_system", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "operating_system");
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

                for (int i = 0; i < rowCount; i++)
                {

                    for (int j = 0; j < AttribCollection.Count; j++)
                    {
                        try
                        {
                            string[] prop = AttribCollection[j].Key.Split('.');

                            if (prop.Length == 1)
                                temp = (threats[i][prop[0]] == null) ? "null" : threats[i][prop[0]].ToString();
                            else if (prop.Length == 2)
                                temp = (threats[i][prop[0]][prop[1]] == null) ? "null" : threats[i][prop[0]][prop[1]].ToString();
                            else if (prop.Length == 3)
                                temp = (threats[i][prop[0]][prop[1]][prop[2]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]].ToString();
                            else if (prop.Length == 4)
                                temp = (threats[i][prop[0]][prop[1]][prop[2]][prop[3]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]][prop[3]].ToString();

                            if (prop[0] == "installed_date")
                            {
                                if (temp == "null")
                                {
                                    temp = "";
                                }
                                else
                                {
                                    DateTime d2 = DateTime.Parse(temp, null, System.Globalization.DateTimeStyles.RoundtripKind);
                                    temp = "=DATE(" + d2.ToString("yyyy,MM,dd") + ")+TIME(" + d2.ToString("HH,mm,ss") + ")";
                                }
                            }

                            dataBlock[i, j] = temp;
                        }
                        catch { }
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
                eHelper.WriteHeaders("Process Data", colCount, rowCount, stopWatch);

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


                formMsg.Message("Creating pivot tables and charts for the process data...", "", allowCancel: false);

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
                MessageBox.Show(ex.Message, "Error extracting Process data", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    pivotWorkSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Process Reports");
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet.Application.DisplayAlerts = false;
                    pivotWorkSheet.Delete();
                    pivotWorkSheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    pivotWorkSheet.Name = "Process Reports";
                    pivotWorkSheet.Activate();
                }
                catch
                {
                    pivotWorkSheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    pivotWorkSheet.Name = "Process Reports";
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
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "Process Reports";
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 1] = "Generated by: " + userName;
                string s = mgmtServer; int start = s.IndexOf("//") + 2; int end = s.IndexOf(".", start); string serverName = s.Substring(start, end - start);
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 1] = "Server: " + serverName;

                Excel.Range rangeStamp = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 11];
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

                // int NextPivotRowDefault = 0;
                int NextPivotRow = 0;
                // int NextChartOffsetDefault = 0;
                int NextChartOffet = 0;

                // Excel.Range rng = null;

                // Chart 1 =========================================================================================
                NextPivotRow = 5;
                NextChartOffet = 80;
                eHelper.CreatePivot("Process Data", colCount, rowCount, "'Process Reports'!R" + NextPivotRow.ToString() + "C11", "PivotTableProcessName", "Process Name", "Process Name", "pid", "Process Count");
                eHelper.CreateChart("PivotTableProcessName", NextChartOffet, "Top 10 Processes");

                #region Chart2
                // Chart 2 =========================================================================================
                /*
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

                eHelper.CreatePivot("Agent Data", colCount, rowCount, "'Agent Reports'!R" + NextPivotRow.ToString() + "C11", "PivotTableOs", "Software Information.Os Name", "Operating System", "ID", "OS Count");
                eHelper.CreateChart("PivotTableOs", NextChartOffet, "Agent Operating Systems");
                */
                #endregion

                #region Chart3
                // Chart 3 =========================================================================================
                /*
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

                eHelper.CreatePivot("Agent Data", colCount, rowCount, "'Agent Reports'!R" + NextPivotRow.ToString() + "C11", "PivotTableAgent", "Agent Version", "Agent Version", "ID", "Agent Count");
                eHelper.CreateChart("PivotTableAgent", NextChartOffet, "Agent Versions");
                */
                #endregion

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
    }
}
