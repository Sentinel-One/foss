using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Forms;
using AboutBoxDemo;
using System.IO;
using System.IO.Compression;
using Excel = Microsoft.Office.Interop.Excel;
using AddinExpress.MSO;
using System.Reflection;
using System.Threading;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace S1ExcelPlugIn
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("FDFBA9E9-6C9D-45B8-8594-C8A035970615"), ProgId("SentinelOneExcelPlugIn.AddinModule")]
    public class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        Crypto crypto = new Crypto();
        private ADXRibbonButton adxRibbonButtonActivity;
        private ADXRibbonSeparator adxRibbonSeparator9;
        private ImageList imageListExcel;
        private ADXRibbonButton adxRibbonButtonAPI;
        private ADXRibbonSeparator adxRibbonSeparator10;
        ExcelHelper eHelper = new ExcelHelper();
        bool SwitchedUser = false;
        private ADXRibbonButton adxRibbonButtonReport;
        private ADXRibbonSeparator adxRibbonSeparator11;
        private ADXRibbonButton adxRibbonButtonApplications;
        private ADXRibbonSeparator adxRibbonSeparator12;
        FormAPI formApiWin = new FormAPI();

        private string token = "";
        private string server = "";
        private string server2 = "";
        private string username = "";
        private int smallDeployment = 100;

        private bool ReferenceDataCached = false;
        FormMessages formMsg = new FormMessages();
        Stopwatch stopWatch = new Stopwatch();
        private ADXRibbonButton adxRibbonButtonProcesses;
        private ADXRibbonSeparator adxRibbonSeparator13;
        bool StartupSesson = true;

        public AddinModule()
        {
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler

        }

        #region Component declarations
        private AddinExpress.XL.ADXExcelTaskPanesManager adxExcelTaskPanesManagerBF;
        private AddinExpress.XL.ADXExcelTaskPanesCollectionItem adxExcelTaskPanesCollectionItemBF;
        private AddinExpress.MSO.ADXRibbonTab adxRibbonTab1;
        private AddinExpress.MSO.ADXRibbonGroup adxRibbonGroup1;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonLogin;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonGroups;
        private ImageList imageListRibbon;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonAgents;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonPDF;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator1;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator2;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator3;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator4;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator5;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonAbout;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonThreats;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonPolicies;
        private AddinExpress.MSO.ADXRibbonLabel adxRibbonLabel1;
        private AddinExpress.MSO.ADXRibbonLabel adxRibbonLabel2;
        private AddinExpress.MSO.ADXRibbonLabel adxRibbonLabel3;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator6;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator7;
        private AddinExpress.MSO.ADXRibbonButton adxRibbonButtonObjects;
        private AddinExpress.MSO.ADXRibbonSeparator adxRibbonSeparator8;
        #endregion

        #region Component Designer generated code
        /// <summary>
        /// Required by designer
        /// </summary>
        private System.ComponentModel.IContainer components;
 
        /// <summary>
        /// Required by designer support - do not modify
        /// the following method
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinModule));
            this.imageListRibbon = new System.Windows.Forms.ImageList(this.components);
            this.adxExcelTaskPanesManagerBF = new AddinExpress.XL.ADXExcelTaskPanesManager(this.components);
            this.adxExcelTaskPanesCollectionItemBF = new AddinExpress.XL.ADXExcelTaskPanesCollectionItem(this.components);
            this.adxRibbonTab1 = new AddinExpress.MSO.ADXRibbonTab(this.components);
            this.adxRibbonGroup1 = new AddinExpress.MSO.ADXRibbonGroup(this.components);
            this.adxRibbonButtonThreats = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.imageListExcel = new System.Windows.Forms.ImageList(this.components);
            this.adxRibbonSeparator1 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonActivity = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator2 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonPolicies = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator3 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonGroups = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator4 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonAgents = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator5 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonApplications = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator6 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonProcesses = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator7 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonObjects = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator8 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonPDF = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator9 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonAPI = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator10 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonReport = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator11 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonLogin = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator12 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonButtonAbout = new AddinExpress.MSO.ADXRibbonButton(this.components);
            this.adxRibbonSeparator13 = new AddinExpress.MSO.ADXRibbonSeparator(this.components);
            this.adxRibbonLabel1 = new AddinExpress.MSO.ADXRibbonLabel(this.components);
            this.adxRibbonLabel2 = new AddinExpress.MSO.ADXRibbonLabel(this.components);
            this.adxRibbonLabel3 = new AddinExpress.MSO.ADXRibbonLabel(this.components);
            // 
            // imageListRibbon
            // 
            this.imageListRibbon.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListRibbon.ImageStream")));
            this.imageListRibbon.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListRibbon.Images.SetKeyName(0, "applications.png");
            this.imageListRibbon.Images.SetKeyName(1, "document_edit.png");
            this.imageListRibbon.Images.SetKeyName(2, "paste.png");
            this.imageListRibbon.Images.SetKeyName(3, "1245774238_Startup Wizard.png");
            this.imageListRibbon.Images.SetKeyName(4, "1245774670_wizard.png");
            this.imageListRibbon.Images.SetKeyName(5, "1245774782_agt_utilities copy.png");
            this.imageListRibbon.Images.SetKeyName(6, "1245775183_news_subscribe.png");
            this.imageListRibbon.Images.SetKeyName(7, "Refresh.png");
            this.imageListRibbon.Images.SetKeyName(8, "programming.png");
            this.imageListRibbon.Images.SetKeyName(9, "wizard.png");
            this.imageListRibbon.Images.SetKeyName(10, "text_editor.png");
            this.imageListRibbon.Images.SetKeyName(11, "package_editors.png");
            this.imageListRibbon.Images.SetKeyName(12, "About.png");
            this.imageListRibbon.Images.SetKeyName(13, "Open.png");
            this.imageListRibbon.Images.SetKeyName(14, "Save.png");
            this.imageListRibbon.Images.SetKeyName(15, "trans_teaser2.png");
            this.imageListRibbon.Images.SetKeyName(16, "export.png");
            this.imageListRibbon.Images.SetKeyName(17, "import.png");
            this.imageListRibbon.Images.SetKeyName(18, "clear.png");
            this.imageListRibbon.Images.SetKeyName(19, "threat-32.png");
            this.imageListRibbon.Images.SetKeyName(20, "policies-32.png");
            this.imageListRibbon.Images.SetKeyName(21, "groups-32.png");
            this.imageListRibbon.Images.SetKeyName(22, "agents-32.png");
            this.imageListRibbon.Images.SetKeyName(23, "devices-32.png");
            this.imageListRibbon.Images.SetKeyName(24, "summary-32.png");
            this.imageListRibbon.Images.SetKeyName(25, "login-32.png");
            this.imageListRibbon.Images.SetKeyName(26, "info-32.png");
            this.imageListRibbon.Images.SetKeyName(27, "about-32.png");
            this.imageListRibbon.Images.SetKeyName(28, "pdf-32.png");
            this.imageListRibbon.Images.SetKeyName(29, "downloads-32.png");
            this.imageListRibbon.Images.SetKeyName(30, "activity-32.png");
            // 
            // adxExcelTaskPanesManagerBF
            // 
            this.adxExcelTaskPanesManagerBF.Items.Add(this.adxExcelTaskPanesCollectionItemBF);
            this.adxExcelTaskPanesManagerBF.SetOwner(this);
            // 
            // adxExcelTaskPanesCollectionItemBF
            // 
            this.adxExcelTaskPanesCollectionItemBF.AlwaysShowHeader = true;
            this.adxExcelTaskPanesCollectionItemBF.CloseButton = true;
            this.adxExcelTaskPanesCollectionItemBF.Enabled = false;
            this.adxExcelTaskPanesCollectionItemBF.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Top;
            this.adxExcelTaskPanesCollectionItemBF.TaskPaneClassName = "BigFixExcelConnector.ADXExcelTaskPaneSessEditor";
            // 
            // adxRibbonTab1
            // 
            this.adxRibbonTab1.Caption = "SentinelOne";
            this.adxRibbonTab1.Controls.Add(this.adxRibbonGroup1);
            this.adxRibbonTab1.Id = "adxRibbonTab_90435f7b6aaa43379be115f0436bc9e5";
            this.adxRibbonTab1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonGroup1
            // 
            this.adxRibbonGroup1.Caption = "SentinelOne PlugIn";
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonThreats);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator1);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonActivity);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator2);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonPolicies);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator3);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonGroups);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator4);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonAgents);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator5);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonApplications);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator6);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonProcesses);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator7);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonObjects);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator8);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonPDF);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator9);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonAPI);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator10);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonReport);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator11);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonLogin);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator12);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonButtonAbout);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonSeparator13);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonLabel1);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonLabel2);
            this.adxRibbonGroup1.Controls.Add(this.adxRibbonLabel3);
            this.adxRibbonGroup1.Id = "adxRibbonGroup_d9db01607ca249cfaa68b209412135ec";
            this.adxRibbonGroup1.ImageList = this.imageListRibbon;
            this.adxRibbonGroup1.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonGroup1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonThreats
            // 
            this.adxRibbonButtonThreats.Caption = "Import Threat Data";
            this.adxRibbonButtonThreats.Description = "Import Threat Data";
            this.adxRibbonButtonThreats.Id = "adxRibbonButton_e26131608b9b4a80ae20e81828105672";
            this.adxRibbonButtonThreats.Image = 8;
            this.adxRibbonButtonThreats.ImageList = this.imageListExcel;
            this.adxRibbonButtonThreats.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonThreats.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonThreats.ScreenTip = "Import Threat Data";
            this.adxRibbonButtonThreats.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonThreats.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonThreats_OnClick);
            // 
            // imageListExcel
            // 
            this.imageListExcel.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListExcel.ImageStream")));
            this.imageListExcel.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListExcel.Images.SetKeyName(0, "icons8-api-32.png");
            this.imageListExcel.Images.SetKeyName(1, "icons8-Info-32.png");
            this.imageListExcel.Images.SetKeyName(2, "icons8-login-32.png");
            this.imageListExcel.Images.SetKeyName(3, "icons8-object-32.png");
            this.imageListExcel.Images.SetKeyName(4, "icons8-agent-32.png");
            this.imageListExcel.Images.SetKeyName(5, "icons8-group-32.png");
            this.imageListExcel.Images.SetKeyName(6, "icons8-policy-32.png");
            this.imageListExcel.Images.SetKeyName(7, "icons8-activity-32.png");
            this.imageListExcel.Images.SetKeyName(8, "icons8-threat-32.png");
            this.imageListExcel.Images.SetKeyName(9, "icons8-pdf-32.png");
            this.imageListExcel.Images.SetKeyName(10, "icons8-Info-purple-32.png");
            this.imageListExcel.Images.SetKeyName(11, "icon8-executive-summary-32.png");
            this.imageListExcel.Images.SetKeyName(12, "icons8-application-32.png");
            this.imageListExcel.Images.SetKeyName(13, "icons8-process-32.png");
            this.imageListExcel.Images.SetKeyName(14, "icons8-disconnected.png");
            this.imageListExcel.Images.SetKeyName(15, "icons8-connected.png");
            // 
            // adxRibbonSeparator1
            // 
            this.adxRibbonSeparator1.Id = "adxRibbonSeparator_97b6badb7f5d49f7b66f5b9583904516";
            this.adxRibbonSeparator1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonActivity
            // 
            this.adxRibbonButtonActivity.Caption = "Import Activity Data";
            this.adxRibbonButtonActivity.Description = "Import Activity Data";
            this.adxRibbonButtonActivity.Id = "adxRibbonButton_0f0f04a2fb91403e8c30ef4fe7d63780";
            this.adxRibbonButtonActivity.Image = 7;
            this.adxRibbonButtonActivity.ImageList = this.imageListExcel;
            this.adxRibbonButtonActivity.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonActivity.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonActivity.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonActivity.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.ImportActivityData_OnClick);
            // 
            // adxRibbonSeparator2
            // 
            this.adxRibbonSeparator2.Id = "adxRibbonSeparator_c9c10bfeca2d4a4b9c50e8ee2c341801";
            this.adxRibbonSeparator2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonPolicies
            // 
            this.adxRibbonButtonPolicies.Caption = "Import Policy Data";
            this.adxRibbonButtonPolicies.Description = "Import Policy Data";
            this.adxRibbonButtonPolicies.Id = "adxRibbonButton_a3bc5b95edf1407f8020ca5bd086203a";
            this.adxRibbonButtonPolicies.Image = 6;
            this.adxRibbonButtonPolicies.ImageList = this.imageListExcel;
            this.adxRibbonButtonPolicies.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonPolicies.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonPolicies.ScreenTip = "Import Policy Data";
            this.adxRibbonButtonPolicies.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonPolicies.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonPolicies_OnClick);
            // 
            // adxRibbonSeparator3
            // 
            this.adxRibbonSeparator3.Id = "adxRibbonSeparator_c89b19b4b16e43ba9ecbc5c108f1548c";
            this.adxRibbonSeparator3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonGroups
            // 
            this.adxRibbonButtonGroups.Caption = "Import Group Data";
            this.adxRibbonButtonGroups.Description = "Import Group Data";
            this.adxRibbonButtonGroups.Id = "adxRibbonButton_58e9597183e349248e8998b6daff8302";
            this.adxRibbonButtonGroups.Image = 5;
            this.adxRibbonButtonGroups.ImageList = this.imageListExcel;
            this.adxRibbonButtonGroups.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonGroups.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonGroups.ScreenTip = "Import Group Data";
            this.adxRibbonButtonGroups.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonGroups.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonGroups_OnClick);
            // 
            // adxRibbonSeparator4
            // 
            this.adxRibbonSeparator4.Id = "adxRibbonSeparator_01f0184d6e6c48a78e40590021b6c081";
            this.adxRibbonSeparator4.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonAgents
            // 
            this.adxRibbonButtonAgents.Caption = "Import Agent Data";
            this.adxRibbonButtonAgents.Description = "Import Agent Data";
            this.adxRibbonButtonAgents.Id = "adxRibbonButton_ec454026ecdb488ea9f9162379c80cce";
            this.adxRibbonButtonAgents.Image = 4;
            this.adxRibbonButtonAgents.ImageList = this.imageListExcel;
            this.adxRibbonButtonAgents.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonAgents.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonAgents.ScreenTip = "Import Agent Data";
            this.adxRibbonButtonAgents.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonAgents.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonAgents_OnClick);
            // 
            // adxRibbonSeparator5
            // 
            this.adxRibbonSeparator5.Id = "adxRibbonSeparator_67e3843ecbcd4d2bb378c652579ae03a";
            this.adxRibbonSeparator5.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonApplications
            // 
            this.adxRibbonButtonApplications.Caption = "Import Application Data";
            this.adxRibbonButtonApplications.Description = "Import Application Data";
            this.adxRibbonButtonApplications.Id = "adxRibbonButton_015385a236714e8eb81d072da70c49c6";
            this.adxRibbonButtonApplications.Image = 12;
            this.adxRibbonButtonApplications.ImageList = this.imageListExcel;
            this.adxRibbonButtonApplications.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonApplications.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonApplications.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonApplications.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonApplications_OnClick);
            // 
            // adxRibbonSeparator6
            // 
            this.adxRibbonSeparator6.Id = "adxRibbonSeparator_c905740a27d2402fb9eccb494bb14008";
            this.adxRibbonSeparator6.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonProcesses
            // 
            this.adxRibbonButtonProcesses.Caption = "Import Process Data";
            this.adxRibbonButtonProcesses.Description = "Import Process Data";
            this.adxRibbonButtonProcesses.Id = "adxRibbonButton_fac704f8e6f948118185425531ae3e4d";
            this.adxRibbonButtonProcesses.Image = 13;
            this.adxRibbonButtonProcesses.ImageList = this.imageListExcel;
            this.adxRibbonButtonProcesses.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonProcesses.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonProcesses.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonProcesses.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonProcesses_OnClick);
            // 
            // adxRibbonSeparator7
            // 
            this.adxRibbonSeparator7.Id = "adxRibbonSeparator_4c5f42c0ea014e68b7eee1def8a68541";
            this.adxRibbonSeparator7.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonObjects
            // 
            this.adxRibbonButtonObjects.Caption = "Import Other Objects";
            this.adxRibbonButtonObjects.Description = "Import Other Objects";
            this.adxRibbonButtonObjects.Id = "adxRibbonButton_f30201fa415d47328b14b4485c3a0705";
            this.adxRibbonButtonObjects.Image = 3;
            this.adxRibbonButtonObjects.ImageList = this.imageListExcel;
            this.adxRibbonButtonObjects.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonObjects.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonObjects.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonObjects.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.ImportAnyObjects_OnClick);
            // 
            // adxRibbonSeparator8
            // 
            this.adxRibbonSeparator8.Id = "adxRibbonSeparator_6708c8411ac04b79bdf5363b1f7d6006";
            this.adxRibbonSeparator8.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonPDF
            // 
            this.adxRibbonButtonPDF.Caption = "Export Sheet to PDF";
            this.adxRibbonButtonPDF.Description = "Export Sheet to PDF";
            this.adxRibbonButtonPDF.Id = "adxRibbonButton_ce810cfbb3c6411c8afdb0d012219451";
            this.adxRibbonButtonPDF.Image = 9;
            this.adxRibbonButtonPDF.ImageList = this.imageListExcel;
            this.adxRibbonButtonPDF.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonPDF.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonPDF.ScreenTip = "Generate Summary Report";
            this.adxRibbonButtonPDF.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonPDF.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.GenerateSummaryReport_OnClick);
            // 
            // adxRibbonSeparator9
            // 
            this.adxRibbonSeparator9.Id = "adxRibbonSeparator_6380c56545524d988a0dcaf80fcefc01";
            this.adxRibbonSeparator9.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonAPI
            // 
            this.adxRibbonButtonAPI.Caption = "Show API Details";
            this.adxRibbonButtonAPI.Description = "Show API Details";
            this.adxRibbonButtonAPI.Id = "adxRibbonButton_3029181719874bab9a8ca3672be98019";
            this.adxRibbonButtonAPI.Image = 0;
            this.adxRibbonButtonAPI.ImageList = this.imageListExcel;
            this.adxRibbonButtonAPI.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonAPI.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonAPI.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonAPI.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.ButtonShowAPI_OnClick);
            // 
            // adxRibbonSeparator10
            // 
            this.adxRibbonSeparator10.Id = "adxRibbonSeparator_c70949500db647cfb0db28e22e127c61";
            this.adxRibbonSeparator10.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonReport
            // 
            this.adxRibbonButtonReport.Caption = "Executive Report";
            this.adxRibbonButtonReport.Id = "adxRibbonButton_99afabcf78c843738ce5fef02c81d2ac";
            this.adxRibbonButtonReport.Image = 11;
            this.adxRibbonButtonReport.ImageList = this.imageListExcel;
            this.adxRibbonButtonReport.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonReport.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonReport.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonReport.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.ExecutiveReport_OnClick);
            // 
            // adxRibbonSeparator11
            // 
            this.adxRibbonSeparator11.Id = "adxRibbonSeparator_76f1ad11f5d943048b58ee8b053be43e";
            this.adxRibbonSeparator11.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonLogin
            // 
            this.adxRibbonButtonLogin.Caption = "Management Server Login";
            this.adxRibbonButtonLogin.Description = "Login to the SentinelOne Management Server";
            this.adxRibbonButtonLogin.Id = "adxRibbonButton_a7012dff158c4e15926ce96980b13f61";
            this.adxRibbonButtonLogin.Image = 2;
            this.adxRibbonButtonLogin.ImageList = this.imageListExcel;
            this.adxRibbonButtonLogin.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonLogin.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonLogin.ScreenTip = "Login to the SentinelOne Management Server";
            this.adxRibbonButtonLogin.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonLogin.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonLogin_OnClick);
            // 
            // adxRibbonSeparator12
            // 
            this.adxRibbonSeparator12.Id = "adxRibbonSeparator_02a2ab61122243cb93f95d1dddec1d43";
            this.adxRibbonSeparator12.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonButtonAbout
            // 
            this.adxRibbonButtonAbout.Caption = "About SentinelOne PlugIn";
            this.adxRibbonButtonAbout.Description = "Some Information About the SentinelOne Excel PlugIn";
            this.adxRibbonButtonAbout.Id = "adxRibbonButton_7954cb173c2443c3b4f781b54e2fcb74";
            this.adxRibbonButtonAbout.Image = 10;
            this.adxRibbonButtonAbout.ImageList = this.imageListExcel;
            this.adxRibbonButtonAbout.ImageTransparentColor = System.Drawing.Color.Transparent;
            this.adxRibbonButtonAbout.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            this.adxRibbonButtonAbout.ScreenTip = "Some Information About the SentinelOne Excel PlugIn";
            this.adxRibbonButtonAbout.Size = AddinExpress.MSO.ADXRibbonXControlSize.Large;
            this.adxRibbonButtonAbout.OnClick += new AddinExpress.MSO.ADXRibbonOnAction_EventHandler(this.adxRibbonButtonAbout_OnClick);
            // 
            // adxRibbonSeparator13
            // 
            this.adxRibbonSeparator13.Id = "adxRibbonSeparator_1026373307614543aedf0edc32dcb686";
            this.adxRibbonSeparator13.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonLabel1
            // 
            this.adxRibbonLabel1.Id = "adxRibbonLabel_a9b48acba1ff4a7499f58ab97ee34646";
            this.adxRibbonLabel1.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonLabel2
            // 
            this.adxRibbonLabel2.Id = "adxRibbonLabel_9548cce728b04ea7828b5c5ace50616a";
            this.adxRibbonLabel2.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // adxRibbonLabel3
            // 
            this.adxRibbonLabel3.Id = "adxRibbonLabel_300d5e25c82145f1ae3d6369d6a7499d";
            this.adxRibbonLabel3.Ribbons = AddinExpress.MSO.ADXRibbons.msrExcelWorkbook;
            // 
            // AddinModule
            // 
            this.AddinName = "SentinelOne Excel PlugIn";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;
            this.AddinInitialize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinInitialize);

        }

        // ================================================================================================
        #endregion

        #region Import Threat Data ====================================================================================
        void adxRibbonButtonThreats_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetThreatData();
        }

        private void GetThreatData()
        {
            try
            {
                FormImportThreats formImportThreatsWin = new FormImportThreats();
                formImportThreatsWin.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Import Activity Data ==================================================================================
        private void ImportActivityData_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetActivityData();
        }

        private void GetActivityData()
        {
            try
            {
                FormImportActivities formImportActivitiesWin = new FormImportActivities();
                formImportActivitiesWin.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Import Policy Data ====================================================================================
        void adxRibbonButtonPolicies_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetPolicyData();
        }

        private void GetPolicyData()
        {
            try
            {
                FormImportAnyObjects formImportObjectsWin = new FormImportAnyObjects();
                formImportObjectsWin.GetData("Policies");
                // formImportPoliciesWin.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Import Group Data =====================================================================================
        void adxRibbonButtonGroups_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetGroupData();
        }

        private void GetGroupData()
        {
            try
            {
                FormImportAnyObjects formImportObjectsWin = new FormImportAnyObjects();
                formImportObjectsWin.GetData("Groups");
                // formImportGroupsWin.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Import Agent Data =====================================================================================
        void adxRibbonButtonAgents_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {

            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetAgentData();
        }

        private void GetAgentData()
        {
            try
            {
                FormImportAgents formImportAgentsWin = new FormImportAgents();
                // formImportAgentsWin.GetData();
                formImportAgentsWin.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Import Application Data ===============================================================================
        private void adxRibbonButtonApplications_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetApplicationData();
        }

        private void GetApplicationData()
        {
            try
            {
                FormImportApplications formImportApplicationsWin = new FormImportApplications();
                formImportApplicationsWin.GetData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Import Process Data =================================================================================
        private void adxRibbonButtonProcesses_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetProcessData();
        }

        private void GetProcessData()
        {
            try
            {
                FormImportProcesses formImportProcessesWin = new FormImportProcesses();
                formImportProcessesWin.GetData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Import Any Objects
        private void ImportAnyObjects_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            if (!LookupTableSheetExists() || !ReferenceDataCached) CacheAllTables();
            if (LookupTableSheetExists() && ReferenceDataCached) GetObjectData();
        }

        private void GetObjectData()
        {
            try
            {
                FormImportAnyObjects formImportObjectsWin = new FormImportAnyObjects();
                // formImportObjectsWin.GetData();
                formImportObjectsWin.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        #endregion

        #region Print
        private void GenerateSummaryReport_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;

            try
            {

                object misValue = System.Reflection.Missing.Value;
                DateTime today = DateTime.Now;
                string fName = "sentinelone_excel_plugin " + today.ToString("yyyy-MM-dd HH-mm-ss") + ".pdf";
                string path = Directory.GetCurrentDirectory();
                string paramExportFilePath = path + "\\" + fName;

                Excel.Worksheet activeWorksheet = ExcelApp.ActiveSheet;

                if (activeWorksheet.Name.Contains("Data"))
                {
                    MessageBox.Show("Please use the normal Excel print function to select how you would like to print the data.\r\nThis button is used to print the sheets with charts and pivot tables.", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                activeWorksheet.PageSetup.TopMargin = 18;
                activeWorksheet.PageSetup.BottomMargin = 18;
                activeWorksheet.PageSetup.LeftMargin = 18;
                activeWorksheet.PageSetup.RightMargin = 18;
                activeWorksheet.PageSetup.FitToPagesWide = 1;
                activeWorksheet.PageSetup.Zoom = false;
                activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                // activeWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                activeWorksheet.PageSetup.CenterHorizontally = true;
                activeWorksheet.PageSetup.CenterVertically = true;

                Excel.XlFixedFormatType paramExportFormat = Excel.XlFixedFormatType.xlTypePDF;
                Excel.XlFixedFormatQuality paramExportQuality = Excel.XlFixedFormatQuality.xlQualityStandard;
                bool paramOpenAfterPublish = true;
                bool paramIncludeDocProps = true;
                bool paramIgnorePrintAreas = true;

                if (ExcelApp.ActiveSheet != null && activeWorksheet.Name.Contains("Reports")) //save as pdf

                    ExcelApp.ActiveSheet.ExportAsFixedFormat(paramExportFormat, paramExportFilePath, paramExportQuality, paramIncludeDocProps, paramIgnorePrintAreas, 1, 1, paramOpenAfterPublish, misValue);

                    /*
                    ExcelApp.ActiveWorkbook.ExportAsFixedFormat(paramExportFormat, paramExportFilePath, paramExportQuality, paramIncludeDocProps, paramIgnorePrintAreas, 1, 2, paramOpenAfterPublish, misValue); */
                   
                else
                    MessageBox.Show("Nothing to export to PDF", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error exporting to PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Show API ==============================================================================================
        private void ButtonShowAPI_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;
            try
            {
                formApiWin.ShowDialog();
                // FormReport frmRpt = new FormReport();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }

        }
        #endregion

        #region Login =================================================================================================
        void adxRibbonButtonLogin_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            ServerLogin();
        }

        private void ServerLogin()
        {
            try
            {
                FormLogin loginWin = new FormLogin();
                loginWin.ShowDialog();

                SwitchedUser = false;
                bool loginSuccess = GetServerInfo();

                if (loginSuccess && SwitchedUser)
                {
                    DeleteOldSheets();
                    ReferenceDataCached = false;
                    CacheAllTables();
                }
                else if (loginSuccess == false)
                {
                    adxRibbonLabel1.Caption = " ";
                    adxRibbonLabel2.Caption = " ";
                    adxRibbonLabel3.Caption = " ";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
            }
        }

        private bool GetServerInfo()
        {
            try
            {
                string apiBatch = crypto.GetSettings("ExcelPlugIn", "ApiBatch");
                if (apiBatch == null)
                {
                    crypto.SetSettings("ExcelPlugIn", "ApiBatch", Globals.ApiBatch.ToString());
                }
                else
                {
                    Globals.ApiBatch = Convert.ToInt32(apiBatch);
                }

                string passphraseBatch = crypto.GetSettings("ExcelPlugIn", "PassphraseBatch");
                if (passphraseBatch == null)
                {
                    crypto.SetSettings("ExcelPlugIn", "PassphraseBatch", Globals.PassphraseBatch.ToString());
                }
                else
                {
                    Globals.PassphraseBatch = Convert.ToInt32(passphraseBatch);
                }

                string applicationBatch = crypto.GetSettings("ExcelPlugIn", "ApplicationBatch");
                if (applicationBatch == null)
                {
                    crypto.SetSettings("ExcelPlugIn", "ApplicationBatch", Globals.ApplicationBatch.ToString());
                }
                else
                {
                    Globals.ApplicationBatch = Convert.ToInt32(applicationBatch);
                }

                token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                server = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                server = server.TrimEnd('/');
                server2 = server.Substring(server.LastIndexOf('/') + 1);
                username = crypto.GetSettings("ExcelPlugIn", "Username");
                string resourceString = server + "/web/api/v1.6/info";
                var restClient = new RestClientInterface(resourceString);
                restClient.Method = HttpVerb.GET;
                restClient.PostData = "";
                var results = restClient.MakeRequest(token).ToString();

                dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(results);
                string version = x.version;
                string build = x.build;
                Globals.ServerVersion = version;

                if (adxRibbonLabel1.Caption != "Logged in as " + username ||
                    adxRibbonLabel2.Caption != "On server " + server2 ||
                    adxRibbonLabel3.Caption != "Version " + version + " " + build
                    )
                {
                    adxRibbonLabel1.Caption = "Logged in as " + username;
                    adxRibbonLabel2.Caption = "On server " + server2;
                    adxRibbonLabel3.Caption = "Version " + version + " " + build;
                    SwitchedUser = true;
                }

                return true;
            }
            catch
            {
                SwitchedUser = false;
                return false;
            }
        }

        private void DeleteOldSheets()
        {
            try
            {
                Excel.Workbook activeWorkBook = null;
                bool foundSheet = false;
                activeWorkBook = (CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook;

                for (int i = 1; i < activeWorkBook.Worksheets.Count; i++)
                {
                    if (activeWorkBook.Worksheets[i].Name.Contains(" Reports") || activeWorkBook.Worksheets[i].Name.Contains(" Data"))
                        foundSheet = true;
                    continue;
                }

                if (foundSheet)
                {
                    DialogResult dialogResult = MessageBox.Show("Do you want to delete all the sheets with data from the previous credentials?", "Please confirm", MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                    else
                    {
                        (CurrentInstance as AddinModule).ExcelApp.ActiveSheet.Application.DisplayAlerts = false;
                        for (int j = activeWorkBook.Worksheets.Count; j > 0; j--)
                        {
                            if (activeWorkBook.Worksheets[j].Name.Contains(" Reports") || activeWorkBook.Worksheets[j].Name.Contains(" Data"))
                            {
                                activeWorkBook.Worksheets[j].Delete();
                                Marshal.ReleaseComObject(activeWorkBook.Worksheets[j]);
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private bool LookupTableSheetExists()
        {
            try
            {
                Excel.Workbook activeWorkBook = null;
                Excel.Worksheet activeWorkSheet = null;
                activeWorkBook = (CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook;
                activeWorkSheet = (Excel.Worksheet)(CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void CacheAllTables()
        {

            try
            {
                if (ReferenceDataCached == true)
                {
                    return;
                }

                Excel.Workbook activeWorkBook = null;
                Excel.Worksheet activeWorkSheet = null;
                activeWorkBook = (CurrentInstance as AddinModule).ExcelApp.ActiveWorkbook;

                if (activeWorkBook == null)
                {
                    return;
                }

                activeWorkSheet = (Excel.Worksheet)(CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Lookup Tables");
                (CurrentInstance as AddinModule).ExcelApp.ActiveSheet.Application.DisplayAlerts = false;
                activeWorkSheet.Delete();

            }
            catch {}

            FindNumberOfAgents();

            CacheLookupTables("groups", 25);

            CacheLookupTables("activities/types", 1);
            CacheLookupTables("agents", 6);
            CacheLookupTables("users", 15);
            CacheLookupTables("policies", 20);

            // activeWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
        }

        private int FindNumberOfAgents()
        {
            string resourceString = server + "/web/api/v1.6/agents/count";
            var restClient = new RestClientInterface(resourceString);
            restClient.Method = HttpVerb.GET;
            var results = restClient.MakeRequest(token).ToString();
            var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
            {
                NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
            };
            dynamic agentResults = Newtonsoft.Json.JsonConvert.DeserializeObject(results, JsonSettings);
            int agentCount = Convert.ToInt32(agentResults.count);
            Globals.TotalAgents = agentCount;
            return agentCount;
        }

        private void CacheLookupTables(string resource, int col)
        {
            Excel.Workbook activeWorkBook = null;
            Excel.Worksheet activeWorkSheet = null;
            try
            {
                try
                {
                    if (ExcelApp.ActiveWorkbook == null)
                    {
                        return;
                    }

                    activeWorkBook = ExcelApp.ActiveWorkbook;
                    activeWorkSheet = (Excel.Worksheet)ExcelApp.Worksheets.get_Item("Lookup Tables");
                    ExcelApp.ActiveSheet.Application.DisplayAlerts = false;
                    activeWorkSheet.Activate();
                }
                catch
                {
                    activeWorkSheet = ExcelApp.ActiveWorkbook.Worksheets.Add() as Excel.Worksheet;
                    activeWorkSheet.Name = "Lookup Tables";
                    activeWorkSheet.Activate();
                }

                activeWorkSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

                string token = crypto.Decrypt(crypto.GetSettings("ExcelPlugIn", "Token"));
                string server = crypto.GetSettings("ExcelPlugIn", "ManagementServer");
                server = server.TrimEnd('/');
                string server2 = server.Substring(server.LastIndexOf('/') + 1);
                string username = crypto.GetSettings("ExcelPlugIn", "Username");

                var JsonSettings = new Newtonsoft.Json.JsonSerializerSettings
                {
                    NullValueHandling = Newtonsoft.Json.NullValueHandling.Include,
                    MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Ignore
                };

                int rowCount = 0;
                int colCount = 5;

                string resourceString = "";
                var restClient = new RestClientInterface();
                restClient.Method = HttpVerb.GET;
                restClient.PostData = "";
                StringBuilder results = new StringBuilder("[");

                if (resource == "agents" && Globals.ServerVersion.Contains("v1.") == false)
                {
                    #region Agents for v2 and above
                    string limit = "limit=" + Globals.ApiBatch.ToString();
                    bool Gogo = true;
                    string last_id = "";
                    int skip_count = 0;
                    dynamic agents = "";
                    dynamic agentsIterate = "";
                    int rowCountTemp = 0;
                    int percentComplete = 0;

                    formMsg.StopProcessing = false;
                    if (!StartupSesson)
                        formMsg.Message("Initializing...", "", allowCancel: false);
                    stopWatch.Start();

                    while (Gogo)
                    {
                        resourceString = server + "/web/api/v1.6/agents/iterator?" + limit + last_id;
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
                        if (Globals.TotalAgents > smallDeployment)
                        {
                            formMsg.Message("Caching data for " + rowCount.ToString("N0") + " of " + Globals.TotalAgents.ToString("N0") + " agents for reference (" + percentComplete.ToString() + "%)...",
                                eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);
                        }
                        if (formMsg.StopProcessing == true)
                        {
                            formMsg.Hide();
                            return;
                        }

                    }

                    stopWatch.Stop();
                    formMsg.Hide();

                    // results.ToString().TrimEnd(',', ' ');
                    results.Replace(",,", "");
                    results.Append("]");

                    ReferenceDataCached = true;
                    #endregion
                }
                else if (resource == "agents" && Globals.ServerVersion.Contains("v1."))
                {
                    #region Agents for v1.x
                    string limit = "limit=" + Globals.ApiBatch.ToString();
                    bool Gogo = true;
                    string skip = "&skip=";
                    int skip_count = 0;
                    dynamic agents = "";
                    int rowCountTemp = 0;
                    int percentComplete = 0;

                    formMsg.StopProcessing = false;
                    if (!StartupSesson)
                        formMsg.Message("Initializing...", "", allowCancel: false);
                    stopWatch.Start();

                    while (Gogo)
                    {
                        // resourceString = server + "/web/api/v1.6/agents?" + limit + skip + skip_count.ToString() + "&include_decommissioned=true";
                        resourceString = server + "/web/api/v1.6/agents?" + limit + skip + skip_count.ToString();
                        restClient.EndPoint = resourceString;
                        restClient.Method = HttpVerb.GET;
                        var res = restClient.MakeRequest(token).ToString();
                        agents = Newtonsoft.Json.JsonConvert.DeserializeObject(res, JsonSettings);
                        rowCountTemp = (int)agents.Count;
                        skip_count = skip_count + Globals.ApiBatch;
                        results.Append(res.TrimStart('[').TrimEnd(']', '\r', '\n')).Append(",");

                        rowCount = rowCount + rowCountTemp;

                        if (rowCountTemp < Globals.ApiBatch)
                            Gogo = false;

                        percentComplete = (int)Math.Round((double)(100 * rowCount) / Globals.TotalAgents);
                        if (Globals.TotalAgents > smallDeployment)
                        {
                            formMsg.Message("Caching data for " + rowCount.ToString("N0") + " of " + Globals.TotalAgents.ToString("N0") + " agents for reference (" + percentComplete.ToString() + "%)...",
                            eHelper.ToReadableStringUpToSec(stopWatch.Elapsed) + " elapsed", allowCancel: true);
                        }
                        if (formMsg.StopProcessing == true)
                        {
                            formMsg.Hide();
                            return;
                        }

                    }

                    stopWatch.Stop();
                    formMsg.Hide();

                    results.ToString().TrimEnd(',', ' ');
                    results.Append("]");

                    ReferenceDataCached = true;
                    #endregion
                }
                else if (resource == "groups")
                {
                    resourceString = server + "/web/api/v1.6/" + resource;
                    restClient.EndPoint = resourceString;
                    var res = restClient.MakeRequest(token).ToString();
                    results.Clear();
                    results.Append(res);
                }
                else
                {
                    resourceString = server + "/web/api/v1.6/" + resource + "?limit=1000000";
                    restClient.EndPoint = resourceString;
                    var res = restClient.MakeRequest(token).ToString();
                    results.Clear();
                    results.Append(res);
                }

                dynamic x = Newtonsoft.Json.JsonConvert.DeserializeObject(results.ToString(), JsonSettings);
                rowCount = (int)x.Count;

                string[,] dataBlock = new string[rowCount, colCount];

                if (resource == "activities/types")
                {
                    for (int i = 0; i < rowCount; i++)
                    {
                        dataBlock[i, 0] = (x[i].id == null) ? "null" : x[i].id.ToString();
                        dataBlock[i, 1] = (x[i].action == null) ? "null" : x[i].action.ToString();
                    }

                    activeWorkSheet.Cells[1, col] = resource;
                    activeWorkSheet.Cells[2, col] = colCount.ToString();
                    activeWorkSheet.Cells[3, col] = rowCount.ToString();
                    activeWorkSheet.Cells[4, col] = "ID";
                    activeWorkSheet.Cells[4, col+1] = "Action";
                }
                else if (resource == "agents")
                {
                    dynamic temp = "";
                    for (int i = 0; i < rowCount; i++)
                    {
                        dataBlock[i, 0] = (x[i].id == null) ? "null" : x[i].id.ToString();
                        dataBlock[i, 1] = (x[i].network_information.computer_name == null) ? "null" : x[i].network_information.computer_name.ToString();
                        dataBlock[i, 2] = (x[i].software_information.os_name == null) ? "[Not Available]" : x[i].software_information.os_name.ToString();

                        dataBlock[i, 3] = (x[i].last_logged_in_user_name == null) ? "[Not Available]" : x[i].last_logged_in_user_name.ToString();
                        dataBlock[i, 4] = (x[i].group_id == null) ? "[Not Available]" : x[i].group_id.ToString();
                        activeWorkSheet.Cells[i+5, col+5].Formula =
                            "=IFERROR(VLOOKUP(" + eHelper.ExcelColumnLetter(9) + (i + 5).ToString() + ",'Lookup Tables'!Y4:Z" + (Globals.GroupRowCount + 4).ToString() + ",2,FALSE),\"[Not Available]\")";
                    }

                    activeWorkSheet.Cells[1, col] = resource;
                    activeWorkSheet.Cells[2, col] = colCount.ToString();
                    activeWorkSheet.Cells[3, col] = rowCount.ToString();
                    activeWorkSheet.Cells[4, col] = "ID";
                    activeWorkSheet.Cells[4, col + 1] = "Network Information.Computer Name";
                    activeWorkSheet.Cells[4, col + 2] = "Software Information.OS Name";
                    activeWorkSheet.Cells[4, col + 3] = "Last Logged In User Name";
                    activeWorkSheet.Cells[4, col + 4] = "Group Id";
                    activeWorkSheet.Cells[4, col + 5] = "Group Name";

                    Globals.AgentRowCount = rowCount;
                }
                else if (resource == "users")
                {
                    for (int i = 0; i < rowCount; i++)
                    {
                        dataBlock[i, 0] = (x[i].id == null) ? "null" : x[i].id.ToString();
                        dataBlock[i, 1] = (x[i].full_name == null) ? "null" : x[i].full_name.ToString();
                    }

                    activeWorkSheet.Cells[1, col] = resource;
                    activeWorkSheet.Cells[2, col] = colCount.ToString();
                    activeWorkSheet.Cells[3, col] = rowCount.ToString();
                    activeWorkSheet.Cells[4, col] = "ID";
                    activeWorkSheet.Cells[4, col + 1] = "Full Name";
                }
                else if (resource == "policies" || resource == "groups")
                {
                    if (resource == "groups")
                        Globals.GroupRowCount = rowCount;

                    for (int i = 0; i < rowCount; i++)
                    {
                        dataBlock[i, 0] = (x[i].id == null) ? "null" : x[i].id.ToString();
                        dataBlock[i, 1] = (x[i].name == null) ? "null" : x[i].name.ToString();
                    }

                    activeWorkSheet.Cells[1, col] = resource;
                    activeWorkSheet.Cells[2, col] = colCount.ToString();
                    activeWorkSheet.Cells[3, col] = rowCount.ToString();
                    activeWorkSheet.Cells[4, col] = "ID";
                    activeWorkSheet.Cells[4, col + 1] = "Name";
                }

                Excel.Range range;
                // range = (CurrentInstance as AddinModule).ExcelApp.get_Range(eHelper.ExcelColumnLetter(col-1) + "5", Missing.Value);
                range = activeWorkSheet.get_Range(eHelper.ExcelColumnLetter(col - 1) + "5", Missing.Value);

                if (x.Count > 0)
                {
                    range = range.Resize[rowCount, colCount];
                    range.Cells.ClearFormats();
                    range.Value = dataBlock;
                    // range.Font.Size = "10";
                }

                if (resource == "activities/types")
                {
                    Excel.Range rangeInt;
                    rangeInt = activeWorkSheet.get_Range(eHelper.ExcelColumnLetter(col - 1) + "5", eHelper.ExcelColumnLetter(col - 1) + (rowCount + 5).ToString());
                    rangeInt.NumberFormat = "0";
                    rangeInt.Cells.HorizontalAlignment = -4152; // Right align the number
                    rangeInt.Value = rangeInt.Value; // Strange technique and workaround to get numbers into Excel. Otherwise, Excel sees them as Text
                }

                activeWorkSheet.Columns.AutoFit();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error caching lookup tables", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (formMsg.Visible)
                {
                    formMsg.Hide();
                }
            }
        }

        private bool CheckIfLoggedIn()
        {
            try
            {
                if (crypto.GetSettings("ExcelPlugIn", "ManagementServer") == null || crypto.GetSettings("ExcelPlugIn", "Token") == null)
                {
                    MessageBox.Show("Please log into the SentinelOne Management Server first", "Please login", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ServerLogin();
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace);
                return false;
            }
        }

        #endregion 

        #region About SentinelOne PlugIn ==============================================================================
        void adxRibbonButtonAbout_OnClick(object sender, AddinExpress.MSO.IRibbonControl control, bool pressed)
        {
            showAboutBox();
        }

        private void showAboutBox()
        {
            AboutBox ab = new AboutBox();
            ab.MoreRichTextBox.Text = "Developed in C# using the SentinelOne Management REST API\n\n";
            ab.ShowDialog();
        }
        #endregion

        #region Executive Report
        private void ExecutiveReport_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            if (!CheckIfLoggedIn()) return;

            Excel.Worksheet threatsWorksheet;
            try
            {
                threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Threat Reports");
                threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Agent Reports");
                threatsWorksheet = (Excel.Worksheet)(ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.Worksheets.get_Item("Application Reports");
            }
            catch
            {
                MessageBox.Show("Please Import all (3) of the Threat, Agent, and Application Reports before running this summary report", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            FormExecutiveSummary formExecSum = new FormExecutiveSummary();
            formExecSum.ShowDialog();
        }
        #endregion

        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }
 
        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        #region Compress and Decompress Routines
        public static void Compress(FileInfo fi)
        {
            // Get the stream of the source file.
            using (FileStream inFile = fi.OpenRead())
            {
                // Prevent compressing hidden and already compressed files.
                if ((File.GetAttributes(fi.FullName) & FileAttributes.Hidden)
                        != FileAttributes.Hidden & fi.Extension != ".gz")
                {
                    // Create the compressed file.
                    using (FileStream outFile = File.Create(fi.FullName + ".gz"))
                    {
                        using (GZipStream Compress = new GZipStream(outFile,
                                CompressionMode.Compress))
                        {
                            // Copy the source file into the compression stream.
                            byte[] buffer = new byte[4096];
                            int numRead;
                            while ((numRead = inFile.Read(buffer, 0, buffer.Length)) != 0)
                            {
                                Compress.Write(buffer, 0, numRead);
                            }
                        }
                    }
                }
            }
        }

        public static void Decompress(FileInfo fi)
        {
            // Get the stream of the source file.
            using (FileStream inFile = fi.OpenRead())
            {
                // Get original file extension, for example "doc" from report.doc.gz.
                string curFile = fi.FullName;
                string origName = curFile.Remove(curFile.Length - fi.Extension.Length);

                //Create the decompressed file.
                using (FileStream outFile = File.Create(origName))
                {
                    using (GZipStream Decompress = new GZipStream(inFile,
                            CompressionMode.Decompress))
                    {
                        //Copy the decompression stream into the output file.
                        byte[] buffer = new byte[4096];
                        int numRead;
                        while ((numRead = Decompress.Read(buffer, 0, buffer.Length)) != 0)
                        {
                            outFile.Write(buffer, 0, numRead);
                        }
                    }
                }
            }
        }
        #endregion

        private void AddinModule_AddinInitialize(object sender, EventArgs e)
        {
            string productVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).ProductVersion;
            string[] versionArray = productVersion.Split('.');
            var newVersion = string.Join(".", versionArray.Take(2));
            adxRibbonButtonAbout.Caption = "About SentinelOne PlugIn " + newVersion;

            Globals.ExcelVersion = ExcelApp.Version;

            // http://www.add-in-express.com/forum/read.php?FID=5&TID=10278
            // The following is set to avoid the Excel: Old format or invalid type library error
            // System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo((AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule).ExcelApp.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI)); 
            if (GetServerInfo())
            {
                if (FindNumberOfAgents() < smallDeployment)
                {
                    CacheAllTables();
                }
                StartupSesson = false;
            }
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }

    }

    public static class Globals
    {
        public static string ExcelVersion = "";
        public static string ServerVersion = "";
        public static int TotalAgents = 0;
        public static int GroupRowCount = 0;
        public static int AgentRowCount = 0;
        public static bool UnresolvedThreatOnly = false;

        public static int ApiBatch = 500;
        public static int PassphraseBatch = 50;
        public static int ApplicationBatch = 20;
        public static int ProcessBatch = 20;

        public static string ApiUrl = "No API calls submitted";
        public static string ApiResults = "No API data available";
        public static int ChartRight = 0;
        public static int ChartBottom = 0;
        public static int PivotBottom = 0;

        public static string ReportPeriod = "";
        public static string StartDate = "";
        public static string EndDate = "";

        public static int TotalThreats = 0;
        public static string ActiveThreats = "";
        public static string ActiveAndUnresolvedThreats = "";
        public static string MostAtRiskGroup = "None";
        public static string MostAtRiskUser = "None";
        public static string MostAtRiskEndpoint = "None";

        // Table
        public static string ThreatData = "";
        public static string DetectionEngine = "";
        public static string InfectedFiles = "";
        public static string MostAtRiskGroups = "";
        public static string MostAtRiskUsers = "";
        public static string MostAtRiskEndpoints = "";

        public static string NetworkStatus = "";
        public static string EndpointOS = "";
        public static string EndpointVersion = "";

        public static string TopApplications = "";

        // Bar Graph Labels
        public static string ThreatDataLabel = "";
        public static string DetectionEnginesLabel = "";
        public static string InfectedFilesLabel = "";
        public static string MostAtRiskGroupsLabel = "";
        public static string MostAtRiskUsersLabel = "";
        public static string MostAtRiskEndpointsLabel = "";

        public static string NetworkStatusLabel = "";
        public static string EndpointOSLabel = "";
        public static string EndpointVersionLabel = "";

        public static string TopApplicationsLabel = "";

        // Bar Graph Values
        public static string ThreatDataValue = "";
        public static string DetectionEnginesValue = "";
        public static string InfectedFilesValue = "";
        public static string MostAtRiskGroupsValue = "";
        public static string MostAtRiskUsersValue = "";
        public static string MostAtRiskEndpointsValue = "";

        public static string NetworkStatusValue = "";
        public static string EndpointOSValue = "";
        public static string EndpointVersionValue = "";

        public static string TopApplicationsValue = "";

    }
}



