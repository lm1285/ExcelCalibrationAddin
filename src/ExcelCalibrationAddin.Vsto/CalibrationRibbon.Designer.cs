namespace ExcelCalibrationAddin.Vsto
{
    partial class CalibrationRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public CalibrationRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tab1 = Factory.CreateRibbonTab();
            this.group1 = Factory.CreateRibbonGroup();
            this.lblAddinVersion = Factory.CreateRibbonLabel();
            this.boxServiceStatus = Factory.CreateRibbonBox();
            this.lblCloudStatus = Factory.CreateRibbonLabel();
            this.lblCloudAddress = Factory.CreateRibbonLabel();
            this.lblCloudIp = Factory.CreateRibbonLabel();
            this.lblCloudDatabase = Factory.CreateRibbonLabel();
            this.lblLoginStatus = Factory.CreateRibbonLabel();
            this.btnQuickGenerate = Factory.CreateRibbonButton();
            this.btnRecognize = Factory.CreateRibbonButton();
            this.btnTemplateLibrary = Factory.CreateRibbonButton();
            this.btnCloudLogin = Factory.CreateRibbonButton();
            this.btnTogglePane = Factory.CreateRibbonButton();
            this.btnSaveSampleData = Factory.CreateRibbonButton();
            this.btnViewSampleData = Factory.CreateRibbonButton();
            this.groupRandomConfig = Factory.CreateRibbonGroup();
            this.boxRandomConfig = Factory.CreateRibbonBox();
            this.lblRandomRangeTitle = Factory.CreateRibbonLabel();
            this.lblRandomRangeDetail = Factory.CreateRibbonLabel();
            this.boxSingleUseOverride = Factory.CreateRibbonBox();
            this.cboOverrideRule = Factory.CreateRibbonComboBox();
            this.edtOverrideRange = Factory.CreateRibbonEditBox();
            this.edtOverrideDecimals = Factory.CreateRibbonEditBox();
            this.btnRandomConfig = Factory.CreateRibbonButton();
            this.groupAlarmValue = Factory.CreateRibbonGroup();
            this.edtAlarmValue = Factory.CreateRibbonEditBox();
            this.groupMultiArea = Factory.CreateRibbonGroup();
            this.btnSaveMultiArea = Factory.CreateRibbonButton();
            this.btnRunMultiArea = Factory.CreateRibbonButton();
            this.btnDeleteMultiArea = Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.boxServiceStatus.SuspendLayout();
            this.groupRandomConfig.SuspendLayout();
            this.boxRandomConfig.SuspendLayout();
            this.boxSingleUseOverride.SuspendLayout();
            this.groupAlarmValue.SuspendLayout();
            this.SuspendLayout();
            //
            // tab1
            //
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabAddIns";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.groupRandomConfig);
            this.tab1.Groups.Add(this.groupAlarmValue);
            this.tab1.Groups.Add(this.groupMultiArea);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.boxServiceStatus);
            this.group1.Items.Add(this.btnQuickGenerate);
            this.group1.Items.Add(this.btnRecognize);
            this.group1.Items.Add(this.btnTemplateLibrary);
            this.group1.Items.Add(this.btnCloudLogin);
            this.group1.Items.Add(this.btnTogglePane);
            this.group1.Items.Add(this.btnSaveSampleData);
            this.group1.Items.Add(this.btnViewSampleData);
            this.group1.Label = "\u6821\u51c6\u52a9\u624b";
            this.group1.Name = "group1";
            // 
            // lblAddinVersion
            // 
            this.lblAddinVersion.Label = "版本 v2026.07.08.6";
            this.lblAddinVersion.Name = "lblAddinVersion";
            //
            // boxServiceStatus
            //
            this.boxServiceStatus.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.boxServiceStatus.Items.Add(this.lblAddinVersion);
            this.boxServiceStatus.Items.Add(this.lblCloudStatus);
            this.boxServiceStatus.Items.Add(this.lblCloudAddress);
            this.boxServiceStatus.Items.Add(this.lblCloudIp);
            this.boxServiceStatus.Items.Add(this.lblCloudDatabase);
            this.boxServiceStatus.Items.Add(this.lblLoginStatus);
            this.boxServiceStatus.Name = "boxServiceStatus";
            //
            // service status labels
            //
            this.lblCloudStatus.Label = "云端：检测中...";
            this.lblCloudStatus.Name = "lblCloudStatus";
            this.lblCloudAddress.Label = "云端地址：";
            this.lblCloudAddress.Name = "lblCloudAddress";
            this.lblCloudIp.Label = "云端IP：检测中...";
            this.lblCloudIp.Name = "lblCloudIp";
            this.lblCloudDatabase.Label = "云数据库：检测中...";
            this.lblCloudDatabase.Name = "lblCloudDatabase";
            this.lblLoginStatus.Label = "登录状态：检测中...";
            this.lblLoginStatus.Name = "lblLoginStatus";
            // 
            // btnQuickGenerate
            // 
            this.btnQuickGenerate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnQuickGenerate.Label = "\u751f\u6210\u968f\u673a\u6570";
            this.btnQuickGenerate.Name = "btnQuickGenerate";
            this.btnQuickGenerate.OfficeImageId = "CalculateNow";
            this.btnQuickGenerate.ScreenTip = "\u6309\u6a21\u677f\u89c4\u5219\u76f4\u63a5\u751f\u6210\u968f\u673a\u6570";
            this.btnQuickGenerate.ShowImage = true;
            this.btnQuickGenerate.SuperTip = "\u8bc6\u522b\u5f53\u524d\u5de5\u4f5c\u7c3f\u5e76\u5339\u914d\u6a21\u677f\u89c4\u5219\uff0c\u7136\u540e\u5c06\u968f\u673a\u6570\u5199\u5165\u6d4b\u91cf\u503c\u533a\u57df\u3002";
            this.btnQuickGenerate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQuickGenerate_Click);
            // 
            // btnRecognize
            // 
            this.btnRecognize.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRecognize.Label = "\u8bc6\u522b\u6a21\u677f";
            this.btnRecognize.Name = "btnRecognize";
            this.btnRecognize.OfficeImageId = "TableDesign";
            this.btnRecognize.ScreenTip = "\u53ea\u626b\u63cf\u6253\u5370\u533a\u57df\u5e76\u8bc6\u522b\u6a21\u677f\u7ed3\u6784";
            this.btnRecognize.ShowImage = true;
            this.btnRecognize.SuperTip = "\u8bfb\u53d6\u5f53\u524d\u5de5\u4f5c\u8868\u6253\u5370\u533a\u57df\u5185\u7684\u8868\u5934\u3001\u533a\u57df\u548c\u683c\u5f0f\u4fe1\u606f\uff0c\u751f\u6210\u8bc6\u522b\u7ed3\u679c\u4e0e\u8349\u7a3f\u89c4\u5219\u3002";
            this.btnRecognize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRecognize_Click);
            // 
            // btnTemplateLibrary
            // 
            this.btnTemplateLibrary.Label = "\u6a21\u677f\u5e93\u7ba1\u7406";
            this.btnTemplateLibrary.Name = "btnTemplateLibrary";
            this.btnTemplateLibrary.OfficeImageId = "FileDocumentManageVersions";
            this.btnTemplateLibrary.ScreenTip = "\u7ba1\u7406\u672c\u5730\u6a21\u677f\u5e93\u7684\u542f\u7528\u4e0e\u5e9f\u6b62\u72b6\u6001";
            this.btnTemplateLibrary.ShowImage = true;
            this.btnTemplateLibrary.SuperTip = "\u5728\u9876\u90e8\u680f\u7ef4\u62a4\u6a21\u677f\u72b6\u6001\uff0c\u8ba9\u5339\u914d\u72b6\u6001\u4e0e\u672c\u5730\u6a21\u677f\u72b6\u6001\u5f62\u6210\u95ed\u73af\u3002";
            this.btnTemplateLibrary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTemplateLibrary_Click);
            //
            // btnCloudLogin
            //
            this.btnCloudLogin.Label = "云端登录";
            this.btnCloudLogin.Name = "btnCloudLogin";
            this.btnCloudLogin.OfficeImageId = "ContactPicture";
            this.btnCloudLogin.ScreenTip = "登录 wzglpt.top 模板服务";
            this.btnCloudLogin.ShowImage = true;
            this.btnCloudLogin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCloudLogin_Click);
            //
            // btnTogglePane
            //
            this.btnTogglePane.Label = "\u4fa7\u8fb9\u680f";
            this.btnTogglePane.Name = "btnTogglePane";
            this.btnTogglePane.OfficeImageId = "NavigationPane";
            this.btnTogglePane.ScreenTip = "\u663e\u793a\u6216\u9690\u85cf\u6821\u51c6\u4fa7\u8fb9\u680f";
            this.btnTogglePane.ShowImage = true;
            this.btnTogglePane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTogglePane_Click);
            // btnSaveSampleData
            this.btnSaveSampleData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSaveSampleData.Label = "保存样本数据";
            this.btnSaveSampleData.Name = "btnSaveSampleData";
            this.btnSaveSampleData.OfficeImageId = "Save";
            this.btnSaveSampleData.ScreenTip = "保存当前模板中的真实测量值样本";
            this.btnSaveSampleData.SuperTip = "请先识别并匹配模板；未匹配时此按钮不可用。";
            this.btnSaveSampleData.ShowImage = true;
            this.btnSaveSampleData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveSampleData_Click);
            // btnViewSampleData
            this.btnViewSampleData.Label = "查看样本数据";
            this.btnViewSampleData.Name = "btnViewSampleData";
            this.btnViewSampleData.OfficeImageId = "ViewDetails";
            this.btnViewSampleData.ShowImage = true;
            this.btnViewSampleData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewSampleData_Click);
            // 
            // groupRandomConfig
            // 
            this.groupRandomConfig.Items.Add(this.boxRandomConfig);
            this.groupRandomConfig.Items.Add(this.boxSingleUseOverride);
            this.groupRandomConfig.Items.Add(this.btnRandomConfig);
            this.groupRandomConfig.Label = "\u968f\u673a\u6570\u914d\u7f6e";
            this.groupRandomConfig.Name = "groupRandomConfig";
            // 
            // boxRandomConfig
            // 
            this.boxRandomConfig.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.boxRandomConfig.Items.Add(this.lblRandomRangeTitle);
            this.boxRandomConfig.Items.Add(this.lblRandomRangeDetail);
            this.boxRandomConfig.Name = "boxRandomConfig";
            // 
            // lblRandomRangeTitle
            // 
            this.lblRandomRangeTitle.Label = "\u793a\u503c\u8bef\u5dee\u968f\u673a\u6570\u751f\u6210\u8303\u56f4";
            this.lblRandomRangeTitle.Name = "lblRandomRangeTitle";
            // 
            // lblRandomRangeDetail
            // 
            this.lblRandomRangeDetail.Label = "\u8bc6\u522b\u6a21\u677f\u540e\u6309 MPE * \u7cfb\u6570\u663e\u793a";
            this.lblRandomRangeDetail.Name = "lblRandomRangeDetail";
            // 
            // boxSingleUseOverride
            // 
            this.boxSingleUseOverride.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.boxSingleUseOverride.Items.Add(this.cboOverrideRule);
            this.boxSingleUseOverride.Items.Add(this.edtOverrideRange);
            this.boxSingleUseOverride.Items.Add(this.edtOverrideDecimals);
            this.boxSingleUseOverride.Name = "boxSingleUseOverride";
            // 
            // cboOverrideRule
            // 
            this.cboOverrideRule.Label = "\u6821\u51c6\u9879";
            this.cboOverrideRule.Name = "cboOverrideRule";
            this.cboOverrideRule.ScreenTip = "\u9009\u62e9\u672c\u6b21\u751f\u6210\u8981\u4e34\u65f6\u8986\u76d6\u7684\u6821\u51c6\u9879";
            this.cboOverrideRule.Text = "";
            // 
            // edtOverrideRange
            // 
            this.edtOverrideRange.Label = "\u7cfb\u6570\u533a\u95f4";
            this.edtOverrideRange.Name = "edtOverrideRange";
            this.edtOverrideRange.ScreenTip = "\u4ec5\u672c\u6b21\u751f\u6210\u751f\u6548\uff0c\u586b MPE \u7cfb\u6570\u533a\u95f4\uff0c\u5982 0.2~0.8 \u6216 \u8d1f:0.2~0.8 \u6b63:0.2~0.9";
            this.edtOverrideRange.Text = "";
            // 
            // edtOverrideDecimals
            // 
            this.edtOverrideDecimals.Label = "\u5c0f\u6570\u4f4d\u6570";
            this.edtOverrideDecimals.Name = "edtOverrideDecimals";
            this.edtOverrideDecimals.ScreenTip = "\u7559\u7a7a\u5219\u4f7f\u7528\u6a21\u677f\u4e2d\u4fdd\u5b58\u7684\u6d4b\u91cf\u503c\u5c0f\u6570\u4f4d\u6570";
            this.edtOverrideDecimals.Text = "";
            // 
            // btnRandomConfig
            // 
            this.btnRandomConfig.Label = "\u914d\u7f6e";
            this.btnRandomConfig.Name = "btnRandomConfig";
            this.btnRandomConfig.OfficeImageId = "DefineName";
            this.btnRandomConfig.ScreenTip = "\u914d\u7f6e\u968f\u673a\u6570\u751f\u6210\u4f9d\u636e";
            this.btnRandomConfig.ShowImage = true;
            this.btnRandomConfig.SuperTip = "\u8bbe\u7f6e\u7ed3\u679c\u533a\u57df\u8ba1\u7b97\u65b9\u6cd5\u3001\u771f\u5b9e\u6d4b\u91cf\u6ce2\u52a8\u3001\u6b63\u6001\u5206\u5e03\u3001\u6807\u51c6\u503c\u53c2\u8003\u548c\u4eea\u5668\u7cbe\u5ea6\u7b49\u751f\u6210\u53c2\u6570\u3002";
            this.btnRandomConfig.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRandomConfig_Click);
            //
            // groupAlarmValue
            //
            this.groupAlarmValue.Items.Add(this.edtAlarmValue);
            this.groupAlarmValue.Label = "\u62a5\u8b66\u503c\u8f93\u5165";
            this.groupAlarmValue.Name = "groupAlarmValue";
            //
            // edtAlarmValue
            //
            this.edtAlarmValue.Label = "\u62a5\u8b66\u503c";
            this.edtAlarmValue.Name = "edtAlarmValue";
            this.edtAlarmValue.ScreenTip = "\u8f93\u5165\u62a5\u8b66\u6821\u51c6\u9879\u672c\u6b21\u8981\u5199\u5165\u7684\u6570\u503c";
            this.edtAlarmValue.SuperTip = "\u70b9\u51fb\u751f\u6210\u968f\u673a\u6570\u65f6\uff0c\u8be5\u6570\u503c\u4f1a\u540c\u65f6\u5199\u5165\u6240\u6709\u62a5\u8b66\u6821\u51c6\u9879\u7684\u6d4b\u91cf\u503c\u533a\u57df\u3002";
            this.edtAlarmValue.Text = "";
            //
            // groupMultiArea
            //
            this.groupMultiArea.Items.Add(this.btnSaveMultiArea);
            this.groupMultiArea.Items.Add(this.btnRunMultiArea);
            this.groupMultiArea.Items.Add(this.btnDeleteMultiArea);
            this.groupMultiArea.Label = "多区域工具";
            this.groupMultiArea.Name = "groupMultiArea";
            //
            // btnSaveMultiArea
            //
            this.btnSaveMultiArea.Label = "保存多区域模板";
            this.btnSaveMultiArea.Name = "btnSaveMultiArea";
            this.btnSaveMultiArea.OfficeImageId = "Save";
            this.btnSaveMultiArea.ScreenTip = "在模板 Excel 中保存不连续区域的位置和尺寸";
            this.btnSaveMultiArea.ShowImage = true;
            this.btnSaveMultiArea.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveMultiArea_Click);
            //
            // btnRunMultiArea
            //
            this.btnRunMultiArea.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRunMultiArea.Label = "运行多区域粘贴";
            this.btnRunMultiArea.Name = "btnRunMultiArea";
            this.btnRunMultiArea.OfficeImageId = "Paste";
            this.btnRunMultiArea.ScreenTip = "匹配已保存模板并将值一键写入当前工作 Excel";
            this.btnRunMultiArea.ShowImage = true;
            this.btnRunMultiArea.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRunMultiArea_Click);
            //
            // btnDeleteMultiArea
            //
            this.btnDeleteMultiArea.Label = "删除多区域模板";
            this.btnDeleteMultiArea.Name = "btnDeleteMultiArea";
            this.btnDeleteMultiArea.OfficeImageId = "Delete";
            this.btnDeleteMultiArea.ScreenTip = "删除本地保存的多区域模板";
            this.btnDeleteMultiArea.ShowImage = true;
            this.btnDeleteMultiArea.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteMultiArea_Click);
            // 
            // CalibrationRibbon
            // 
            this.Name = "CalibrationRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CalibrationRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.boxServiceStatus.ResumeLayout(false);
            this.boxServiceStatus.PerformLayout();
            this.groupRandomConfig.ResumeLayout(false);
            this.groupRandomConfig.PerformLayout();
            this.boxRandomConfig.ResumeLayout(false);
            this.boxRandomConfig.PerformLayout();
            this.boxSingleUseOverride.ResumeLayout(false);
            this.boxSingleUseOverride.PerformLayout();
            this.groupAlarmValue.ResumeLayout(false);
            this.groupAlarmValue.PerformLayout();
            this.groupMultiArea.ResumeLayout(false);
            this.groupMultiArea.PerformLayout();
            this.ResumeLayout(false);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblAddinVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxServiceStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblCloudStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblCloudAddress;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblCloudIp;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblCloudDatabase;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblLoginStatus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQuickGenerate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRecognize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTemplateLibrary;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCloudLogin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTogglePane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveSampleData;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnViewSampleData;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupRandomConfig;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxRandomConfig;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblRandomRangeTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblRandomRangeDetail;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox boxSingleUseOverride;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cboOverrideRule;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edtOverrideRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edtOverrideDecimals;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRandomConfig;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAlarmValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox edtAlarmValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMultiArea;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveMultiArea;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRunMultiArea;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteMultiArea;
    }

    partial class ThisRibbonCollection
    {
        internal CalibrationRibbon CalibrationRibbon
        {
            get { return this.GetRibbon<CalibrationRibbon>(); }
        }
    }
}
