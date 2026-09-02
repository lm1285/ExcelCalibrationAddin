using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Core.Models;
using ExcelCalibrationAddin.Core.Services;
using ExcelCalibrationAddin.Host.ViewModels;
using ExcelCalibrationAddin.Host.Vsto;
using ExcelCalibrationAddin.Vsto.TaskPane;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCalibrationAddin.Vsto
{
    public partial class ThisAddIn
    {
        private const string TaskPaneTitle = "校准助手";

        private VstoAddinFacade _facade;
        private readonly object _facadeInitializationLock = new object();
        private string _configPath;
        private GenerationConfiguration _generationConfiguration;
        private TaskPane.CalibrationTaskPaneControl _taskPaneControl;
        private CustomTaskPane _taskPane;
        private bool _restoringTaskPaneVisibility;
        private bool _isNavigatingFromPlugin;
        private DateTime _lastProgressUiUpdateUtc = DateTime.MinValue;
        private int _lastProgressPercent = -1;
        private string _lastProgressMessage = string.Empty;
        private CellRange _lastHighlightedRange;
        private string _lastMatchedWorkbookKey = string.Empty;
        private TaskPaneState _lastGenerationState;
        private Task _facadeWarmupTask = Task.CompletedTask;
        private readonly object _workbookMatchTaskLock = new object();
        private Task _activeWorkbookMatchTask = Task.CompletedTask;
        private string _activeWorkbookMatchKey = string.Empty;
        private Task _startupTemplateSyncTask = Task.CompletedTask;
        private bool _isWritingGeneratedValues;
        private DailySyncScheduler _dailySyncScheduler;
        private YingdaoAutomationServer _yingdaoAutomationServer;
        private Process _yingdaoAutomationBridge;
        private SynchronizationContext _excelUiSynchronizationContext;
        private System.Threading.Timer _serviceStatusTimer;

        internal void ToggleTaskPane()
        {
            EnsureTaskPane();
            var visible = !IsTaskPaneVisible();
            if (!visible && !_taskPaneControl.ConfirmCloseIfDirty())
            {
                return;
            }
            SetTaskPaneVisible(visible);
            this.Application.StatusBar = visible ? "校准助手：侧边栏已打开" : "校准助手：侧边栏已隐藏";
            Trace.WriteLine($"[VSTO] Task pane visible={IsTaskPaneVisible()}");
        }

        internal ExcelCalibrationAddin.Host.ViewModels.TaskPaneState GetCachedGenerationState() => _lastGenerationState;

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
