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
using System.Data.Entity.Design.PluralizationServices;
using System.Diagnostics;
using System.Text;

namespace S1ExcelPlugIn
{
    public partial class FormImportAnyObjects : Form
    {
        #region Class variables
        Crypto crypto = new Crypto();

        int colCount = 0;
        int rowCount = 0;
        string userName = "";

        /*
        int r_color = 104;
        int g_color = 33;
        int b_color = 122; */

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

        bool loggedIn = false;

        #endregion

        public FormImportAnyObjects()
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
                else
                {
                    loggedIn = true;
                }

                if (Globals.ServerVersion.StartsWith("v1."))
                {
                    for (int n = listBoxObjects.Items.Count - 1; n >= 0; --n)
                    {
                        string removelistitem = "Agents/Configuration";
                        if (listBoxObjects.Items[n].ToString().Contains(removelistitem))
                        {
                            listBoxObjects.Items.RemoveAt(n);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error initialzing saved parameters, no harm done", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Button events
        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void buttonImportData_Click(object sender, EventArgs e)
        {
            GetData(listBoxObjects.SelectedItem.ToString());
            Close();
        }

        private void buttonImportAndGenerate_Click(object sender, EventArgs e)
        {
            GetData(listBoxObjects.SelectedItem.ToString());
            GenerateReport();
            GenerateChart();
            Close();
        }

        #endregion 

        public void GetData(string selectedObject)
        {
            try
            {
                if (!loggedIn)
                {
                    return;
                }

                #region Create worksheet

                CultureInfo ci = new CultureInfo("en-us");
                PluralizationService ps = PluralizationService.CreateService(ci);
                string selectedObjSingular = ps.Singularize(selectedObject);

                // Creates a worksheet for "Threats" if one does not already exist
                // =================================================================================================================
                Excel.Worksheet threatsWorksheet;
                try
                {
                    threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item(selectedObjSingular.Replace('/', ' ') + " Data");
                    threatsWorksheet.Activate();
                }
                catch
                {
                    threatsWorksheet = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    threatsWorksheet.Name = selectedObjSingular.Replace('/',' ') + " Data";
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

                string limit = "limit=" + Globals.ApiBatch.ToString();
                bool Gogo = true;
                string skip = "&skip=";
                int skip_count = 0;
                StringBuilder results = new StringBuilder("[");
                int maxColumnWidth = 80;
                dynamic res = "";
                int rowCountTemp = 0;
                stopWatch.Start();

                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };

                formMsg.StopProcessing = false;
                formMsg.StartMessage();

                this.Opacity = 0;
                this.ShowInTaskbar = false;

                string resourceString = "";

                var restClient = new RestClientInterface();

                #region Agent Passphrase
                if (selectedObject == "Agents/Passphrase" || selectedObject == "Agents/Configuration")
                {
                    #region Find All the Agents from the hidden Lookup Table sheet
                    Excel.Worksheet lookupSheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                    int rows = Convert.ToInt32(lookupSheet.Cells[3, 6].Value) + 4;
                    Excel.Range AgentIDs = lookupSheet.Range["F5:H" + rows.ToString()];
                    System.Array AgentArray;
                    int AgentArrayItems = 0;
                    AgentArray = AgentIDs.Value;
                    AgentArrayItems = AgentArray.GetLength(0);
                    int UpdateInterval = Globals.PassphraseBatch;
                    #endregion

                    for (int i = 1; i <= AgentArrayItems; i++)
                    {
                        // resourceString = mgmtServer + "/web/api/v1.6/agents/" + AgentArray.GetValue(i, 1).ToString() + "/passphrase";
                        resourceString = mgmtServer + "/web/api/v1.6/agents/" + AgentArray.GetValue(i, 1).ToString() + selectedObject.ToLower().Substring(selectedObject.IndexOf("/"));
                        restClient.EndPoint = resourceString;
                        restClient.Method = HttpVerb.GET;
                        var batch_string = restClient.MakeRequest(token).ToString();
                        batch_string = batch_string.Replace("\"configuration\": \"", "\"configuration\": ");
                        batch_string = batch_string.Replace("\\\"", "\"");
                        batch_string = batch_string.Replace("\\\\\\\\", "\\\\");
                        batch_string = batch_string.Replace("}\", \"updated_at\"", "}, \"updated_at\"");

                        int percentComplete = 0;

                        if (i % UpdateInterval == 0)
                        {
                            percentComplete = (int)Math.Round((double)(100 * i) / AgentArrayItems);
                            formMsg.UpdateMessage("Collecting " + selectedObject.ToLower().Substring(selectedObject.IndexOf("/") + 1) + " from " + i.ToString("N0") + " of " + AgentArrayItems.ToString("N0") + " agents (" + percentComplete.ToString() + "%)...",
                                eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed");
                            if (formMsg.StopProcessing == true)
                            {
                                formMsg.Hide();
                                return;
                            }
                        }

                        JObject oneAgent = JObject.Parse(batch_string);

                        var PropertyComputerName = new JProperty("computer_name", AgentArray.GetValue(i, 2) == null ? "N/A" : AgentArray.GetValue(i, 2).ToString());
                        var PropertyOperatingSystem = new JProperty("operating_system", AgentArray.GetValue(i, 3) == null ? "N/A" : AgentArray.GetValue(i, 3).ToString());

                        oneAgent.AddFirst(PropertyOperatingSystem);
                        oneAgent.AddFirst(PropertyComputerName);

                        // oneAgent.Add("operating_system", AgentArray.GetValue(i, 3) == null ? "N/A" : AgentArray.GetValue(i, 3).ToString());
                        // oneAgent.Add("computer_name", AgentArray.GetValue(i, 2) == null ? "N/A" : AgentArray.GetValue(i, 2).ToString());

                        results.Append(oneAgent.ToString()).Append(",");
                        rowCount++;
                    }

                    // MessageBox.Show(results.ToString());
                }
                #endregion

                else if (selectedObject == "Groups")
                {
                    resourceString = mgmtServer + "/web/api/v1.6/" + selectedObject.ToLower();
                    restClient.EndPoint = resourceString;
                    restClient.Method = HttpVerb.GET;
                    var batch_string = restClient.MakeRequest(token).ToString();
                    res = Newtonsoft.Json.JsonConvert.DeserializeObject(batch_string, JsonSettings);
                    rowCountTemp = (int)res.Count;
                    skip_count = skip_count + Globals.ApiBatch;
                    results.Append(batch_string.TrimStart('[').TrimEnd(']', '\r', '\n')).Append(",");

                    rowCount = rowCount + rowCountTemp;
                }

                #region All other objects
                else
                {
                    while (Gogo)
                    {
                        resourceString = mgmtServer + "/web/api/v1.6/" + selectedObject.ToLower() + "?" + limit + skip + skip_count.ToString();
                        restClient.EndPoint = resourceString;
                        restClient.Method = HttpVerb.GET;
                        var batch_string = restClient.MakeRequest(token).ToString();
                        res = Newtonsoft.Json.JsonConvert.DeserializeObject(batch_string, JsonSettings);
                        rowCountTemp = (int)res.Count;
                        skip_count = skip_count + Globals.ApiBatch;
                        results.Append(batch_string.TrimStart('[').TrimEnd(']', '\r', '\n')).Append(",");

                        rowCount = rowCount + rowCountTemp;

                        if (rowCountTemp < Globals.ApiBatch)
                            Gogo = false;

                        formMsg.UpdateMessage("Loading " + selectedObject.ToLower() + " data: " + rowCount.ToString(),
                            eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed");

                        if (formMsg.StopProcessing == true)
                        {
                            formMsg.Hide();
                            return;
                        }

                    }
                }
                #endregion

                stopWatch.Stop();
                formMsg.Hide();

                Globals.ApiUrl = resourceString;

                // results.ToString().TrimEnd(',', ' ');
                results.Length--;
                results.Append("]");

                // MessageBox.Show(results.ToString());

                if (results.ToString() == "[,]")
                {
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "No " + selectedObject + " data found";
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Select();
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Columns.AutoFit();
                    formMsg.Hide();
                    return;
                }

                dynamic tempItem = Newtonsoft.Json.JsonConvert.DeserializeObject(results.ToString(), JsonSettings);
                // =======================================================================================

                Newtonsoft.Json.Linq.JArray threats = new JArray();

                if (tempItem.GetType() == typeof(Newtonsoft.Json.Linq.JObject))
                {
                    threats.Add(tempItem);
                    results.Append("]");
                }
                else
                {
                    threats = tempItem;
                }

                rowCount = (int)threats.Count;

                #endregion

                #region Parse attribute headers
                JArray ja = JArray.Parse(results.ToString());

                // This is saved for viewing in the Show API window
                Globals.ApiResults = ja.ToString();

                // Stop processing if no data found
                if (ja.Count == 0)
                {
                    (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "No " + selectedObject + " data found";
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
                    else if (dataType == "Array" && Level1.Key == "engine_data")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "engine", "String"));
                        AttribNo++;
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "asset_name", "String"));
                        AttribNo++;
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key + "." + "asset_version", "String"));
                        AttribNo++;
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
                    else if (Level1.Key == "policy_id")
                    {
                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>(Level1.Key, dataType));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, dataType, Level1.Key);

                        AttribCollection.Insert(AttribNo, new KeyValuePair<string, string>("policy_name", "Reference"));
                        AttribNo++;
                        AddFormatter(AttribNo - 1, "Reference", "policy_name");
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
                            {
                                if (threats[i][prop[0]][prop[1]] != null && threats[i][prop[0]][prop[1]][prop[2]] != null)
                                {
                                    temp = (threats[i][prop[0]][prop[1]][prop[2]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]].ToString();
                                }
                                else
                                {
                                    temp = "null";
                                }
                            }
                            else if (prop.Length == 4)
                                temp = (threats[i][prop[0]][prop[1]][prop[2]][prop[3]] == null) ? "null" : threats[i][prop[0]][prop[1]][prop[2]][prop[3]].ToString();

                            if (prop[0] == "created_date" || prop[0] == "meta_data" || prop[0] == "last_active_date")
                            {
                                DateTime d2 = DateTime.Parse(temp, null, System.Globalization.DateTimeStyles.RoundtripKind);
                                temp = "=DATE(" + d2.ToString("yyyy,MM,dd") + ")+TIME(" + d2.ToString("HH,mm,ss") + ")";
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
                eHelper.WriteHeaders(selectedObjSingular.Replace('/', ' ') + " Data", colCount, rowCount, stopWatch);

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


            }
            catch (Exception ex)
            {
                formMsg.Hide();
                MessageBox.Show(ex.Message, "Error extracting data", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                Excel.Range titleRow = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A1", "CA1");
                titleRow.Select();
                titleRow.RowHeight = 33;
                titleRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color));
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[1, 1] = "Threat Reports";
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[2, 1] = "Generated by: " + userName;
                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Cells[3, 1] = DateTime.Now.ToString("f");

                Excel.Range rowSeparator = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A3", "CA3");
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r_color, g_color, b_color)); // 
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = 1; // xlContinuous
                rowSeparator.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = 4; // xlThick
                #endregion

                // Create the Pivot Table
                pivotCaches = activeWorkBook.PivotCaches();
                activeWorkBook.ShowPivotTableFieldList = false;
                // pivotCache = pivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "Threats!$A$4:$" + ExcelColumnLetter(colCount) + "$" + rowCount);
                // string rangeName = "Threats!$A$4:$T$100";
                string rangeName = "'Agent Data'!$A$4:$" + eHelper.ExcelColumnLetter(colCount-1) + "$" + (rowCount + 4).ToString();
                pivotCache = pivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, rangeName);
                // pivotTable = pivotCache.CreatePivotTable("Reports!R3C1");
                pivotTable = pivotCache.CreatePivotTable("'Agent Reports'!R7C1");
                pivotTable.NullString = "0";

                // Set the Pivot Fields
                pivotFields = (Excel.PivotFields)pivotTable.PivotFields();

                // Month Pivot Field
                monthPivotField = (Excel.PivotField)pivotFields.Item("Created Date");
                monthPivotField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                monthPivotField.Position = 1;
                monthPivotField.DataRange.Cells[1].Group(true, true, Type.Missing, new bool[] { false, false, false, false, true, true, true });

                // Mitigation Status Pivot Field
                statusPivotField = (Excel.PivotField)pivotFields.Item("Mitigation Status");
                statusPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

                // Resolved Pivot Field
                resolvedPivotField = (Excel.PivotField)pivotFields.Item("Resolved");
                resolvedPivotField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;

                // Threat ID Pivot Field
                threatIdPivotField = (Excel.PivotField)pivotFields.Item("ID");

                // Count of Threat ID Field
                threatIdCountPivotField = pivotTable.AddDataField(threatIdPivotField, "# of Threats", Excel.XlConsolidationFunction.xlCount);

                slicerCaches = activeWorkBook.SlicerCaches;
                // Month Slicer
                monthSlicerCache = slicerCaches.Add(pivotTable, "Created Date", "CreatedDate");
                monthSlicers = monthSlicerCache.Slicers;
                monthSlicer = monthSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Created Date", "Created Date", 80, 480, 144, 100);
                // Mitigation Status Slicer
                statusSlicerCache = slicerCaches.Add(pivotTable, "Mitigation Status", "MitigationStatus");
                statusSlicers = statusSlicerCache.Slicers;
                statusSlicer = statusSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Mitigation Status", "Mitigation Status", 80, 634, 144, 100);
                // Resolved Slicer
                resolvedSlicerCache = slicerCaches.Add(pivotTable, "Resolved", "Resolved");
                resolvedSlicers = resolvedSlicerCache.Slicers;
                resolvedSlicer = resolvedSlicers.Add((ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveSheet, Type.Missing, "Resolved", "Resolved", 80, 788, 144, 100);
                // Slicer original sizes top 15, width 144, height 200
            }
            catch (Exception ex)
            {
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
                // selectedRange = (Excel.Range)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Selection;
                selectedRange = IdentifyPivotRanges(rangePivot);
                shapes = activeSheet.Shapes;
                shapes.AddChart2(Style: 201, XlChartType: Excel.XlChartType.xlColumnClustered,
                    Left: 480, Top: 190, Width: 450,
                    Height: Type.Missing, NewLayout: true).Select();

                chart = (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.ActiveChart;
                chart.HasTitle = false;
                // chart.ChartTitle.Text = "Threats History";
                chart.ChartArea.Interior.Color = System.Drawing.Color.FromArgb(242, 244, 244); // Change chart to light gray
                // chart.ChartArea.Interior.Color = System.Drawing.Color.FromRgb(0, 255, 0);
                chart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowValue, true, true, true, true, true, true, true, true, true); // Turn on data labels
                chart.HasLegend = true;
                chart.SetSourceData(selectedRange);

                (ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.get_Range("A2", "A2").Select();
            }
            finally
            {
                if (chartTitle != null) Marshal.ReleaseComObject(chartTitle);
                if (chart != null) Marshal.ReleaseComObject(chart);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (selectedRange != null) Marshal.ReleaseComObject(selectedRange);
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

        private Excel.Range IdentifyPivotRanges(Excel.Range xlRange)
        {
            Excel.Range pivotRanges = null;
            Excel.PivotTables pivotTables = xlRange.Worksheet.PivotTables();
            int pivotCount = pivotTables.Count;
            for (int i = 1; i <= pivotCount; i++)
            {
                Excel.Range tmpRange = xlRange.Worksheet.PivotTables(i).TableRange2;
                if (pivotRanges == null) pivotRanges = tmpRange;
                // pivotRanges = this.Application.Union(pivotRanges, tmpRange);
            }
            return pivotRanges;
        }


        #endregion

    }
}
