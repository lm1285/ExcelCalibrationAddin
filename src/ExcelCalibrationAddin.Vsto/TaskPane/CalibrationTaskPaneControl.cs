using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelCalibrationAddin.Contracts;
using ExcelCalibrationAddin.Host.ViewModels;

namespace ExcelCalibrationAddin.Vsto.TaskPane
{
    public partial class CalibrationTaskPaneControl : UserControl
    {
	private static readonly IntPtr PerMonitorV2DpiContext = new IntPtr(-4);

	[DllImport("user32.dll", EntryPoint = "SetThreadDpiAwarenessContext", SetLastError = true)]
	private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

	private sealed class FieldStatusTag
	{
		public int RowIndex { get; }

		public int ColumnIndex { get; }

		public string ColumnName { get; }

		public FieldStatusTag(int rowIndex, int columnIndex, string columnName)
		{
			RowIndex = rowIndex;
			ColumnIndex = columnIndex;
			ColumnName = columnName;
		}
	}

	private sealed class MappingCardLayout
	{
		public int RowIndex { get; set; }
		public Panel Card { get; set; }
		public Control StandardValuePanel { get; set; }
		public TableLayoutPanel FieldGrid { get; set; }
		public Label StatusLabel { get; set; }
		public bool IsCollapsed { get; set; }
	}

	private sealed class ExcelToggleSwitch : CheckBox
	{
		public ExcelToggleSwitch()
		{
			AutoSize = false;
			Text = string.Empty;
			Cursor = Cursors.Hand;
			SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
		}

		protected override void OnPaintBackground(PaintEventArgs e)
		{
		}

		protected override void OnPaint(PaintEventArgs e)
		{
			e.Graphics.Clear(Parent?.BackColor ?? Color.White);
			e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
			e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
			e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
			int trackHeight = Math.Max(14, Height - 4);
			int trackWidth = Math.Max(trackHeight + 8, Width - 2);
			int top = (Height - trackHeight) / 2;
			Color trackColor = Checked ? Color.FromArgb(33, 115, 70) : Color.FromArgb(198, 200, 204);
			using (Brush brush = new SolidBrush(trackColor))
			{
				e.Graphics.FillEllipse(brush, 1, top, trackHeight, trackHeight);
				e.Graphics.FillRectangle(brush, 1 + trackHeight / 2, top, trackWidth - trackHeight, trackHeight);
				e.Graphics.FillEllipse(brush, trackWidth - trackHeight - 1, top, trackHeight, trackHeight);
			}

			int thumbSize = trackHeight - 4;
			int thumbLeft = Checked ? trackWidth - thumbSize - 3 : 3;
			using (Brush brush = new SolidBrush(Color.White))
			{
				e.Graphics.FillEllipse(brush, thumbLeft, top + 2, thumbSize, thumbSize);
			}
		}

		protected override void OnCheckedChanged(EventArgs e)
		{
			base.OnCheckedChanged(e);
			Invalidate();
		}
	}

	private const string ColumnProject = "Project";

	private const string ColumnSection = "Section";
	private const string ColumnSetpoint = "Setpoint";

	private const string ColumnStandard = "Standard";

	private const string ColumnMeasurement = "Measurement";

	private const string ColumnAverage = "Average";

	private const string ColumnError = "Error";

	private const string ColumnRequirement = "Requirement";

	private const string ColumnUncertainty = "Uncertainty";

	private const string ColumnRange = "Range";

	private const string ColumnResult = "Result";

	private readonly Label _progressLabel = new Label();

	private readonly ProgressBar _progressBar = new ProgressBar();

	private readonly Label _matchChainLabel = new Label();

	private readonly Panel _overviewCard = new Panel();

	private readonly Label _statusIcon = new Label();

	private readonly Label _statusSummary = new Label();

	private readonly Button _saveTemplateButton = CreateActionButton("保存模板");
	private readonly Button _saveAsTemplateButton = CreateActionButton("另存为");

	private readonly Button _addCalibrationItemButton = CreateActionButton("新增校准项");

	private readonly Button _deleteSelectedCalibrationItemsButton = CreateActionButton("删除");

	private readonly Button _useSelectionForFieldButton = CreateActionButton("设为当前选区");

	private readonly Button _clearSelectedFieldRangeButton = CreateActionButton("清除区域");

	private readonly Label _fieldSelectionLabel = new Label();

	private readonly Panel _fieldActionBar = new Panel();


	private readonly FlowLayoutPanel _mappingCards = new FlowLayoutPanel();

	private readonly Label _emptyLabel = new Label();

	private readonly List<Control> _fieldStatusCards = new List<Control>();

	private readonly HashSet<int> _manualStandardRows = new HashSet<int>();

	private readonly Dictionary<int, double?> _autoStandardValues = new Dictionary<int, double?>();


	private readonly HashSet<int> _selectedCalibrationRows = new HashSet<int>();

	private readonly HashSet<int> _collapsedCalibrationRows = new HashSet<int>();

	private bool _hasInitializedCollapsedCards;

	private TemplateFingerprint _currentFingerprint;

	private List<MeasurementRule> _currentRules = new List<MeasurementRule>();

	private List<TemplateRegionMapping> _currentMappings = new List<TemplateRegionMapping>();

	private ExcelCalibrationAddin.Core.Models.GenerationConfiguration _appliedGenerationConfiguration;

	private bool _usesTemplateGenerationConfiguration;

	private bool _featuresBlocked;

	private bool _stateFeaturesBlocked;

	private bool _canGenerate;

	private bool? _preferredCreateNewTemplate;

	private bool _readOnlyMode;
	private bool _isBinding;
	private bool _hasUnsavedChanges;
	private bool _isEditingSavedTemplate;
	private string _editingTemplateFingerprint = string.Empty;
	private string _editingRemoteTemplateId = string.Empty;
	private string _editingTemplateName = string.Empty;
	private TemplateDirectoryMetadata _editingDirectoryMetadata;

	private int _highlightedRowIndex = -1;

	private int _highlightedColumnIndex = -1;


	public CalibrationTaskPaneControl()
	{
		AutoScaleMode = AutoScaleMode.Dpi;
		AutoScaleDimensions = new SizeF(96f, 96f);
		// Keep the fixed-width editor usable when Excel gives the task pane a
		// smaller client area (for example on a high-DPI laptop display). The
		// parent panel can scroll instead of shrinking buttons and field cards.
		MinimumSize = new Size(380, 0);
		DoubleBuffered = true;
		ResizeRedraw = true;
		IntPtr previousDpiContext = TrySetPerMonitorV2DpiContext();
		try
		{
			InitializeComponent();
			BuildLayout();
		}
		finally
		{
			RestoreDpiContext(previousDpiContext);
		}
	}

	private static IntPtr TrySetPerMonitorV2DpiContext()
	{
		try
		{
			return SetThreadDpiAwarenessContext(PerMonitorV2DpiContext);
		}
		catch (EntryPointNotFoundException)
		{
			return IntPtr.Zero;
		}
	}

	private static void RestoreDpiContext(IntPtr previousDpiContext)
	{
		if (previousDpiContext == IntPtr.Zero)
		{
			return;
		}

		try
		{
			SetThreadDpiAwarenessContext(previousDpiContext);
		}
		catch (EntryPointNotFoundException)
		{
		}
	}


	public void Bind(TaskPaneState state)
	{
		if (base.InvokeRequired)
		{
			BeginInvoke((MethodInvoker)delegate
			{
				Bind(state);
			});
			return;
		}
		state = (TaskPaneState)(((object)state) ?? ((object)new TaskPaneState()));
		_isBinding = true;
		_isEditingSavedTemplate = false;
		_editingTemplateFingerprint = string.Empty;
		_editingRemoteTemplateId = string.Empty;
		_editingTemplateName = string.Empty;
		_editingDirectoryMetadata = null;
		_hasUnsavedChanges = false;
		_manualStandardRows.Clear();
		_autoStandardValues.Clear();
		_selectedCalibrationRows.Clear();
		_currentFingerprint = TaskPaneModelCloner.CloneFingerprint(state.Fingerprint);
		_currentRules = TaskPaneModelCloner.CloneRules(state.DraftRules);
		_currentMappings = TaskPaneModelCloner.CloneMappings(state.MappingItems);
		AlignRulesToMappings();
		for (int rowIndex = 0; rowIndex < _currentRules.Count; rowIndex++)
		{
			if (_currentRules[rowIndex]?.ManualStandardValues?.Count > 0)
			{
				_manualStandardRows.Add(rowIndex);
			}
		}
		_appliedGenerationConfiguration = CloneGenerationConfiguration(state.AppliedGenerationConfiguration);
		_usesTemplateGenerationConfiguration = state.UsesTemplateGenerationConfiguration;
		if (!state.CanGenerate)
		{
			AlignRulesToMappings();
		}
		CaptureAutoStandardValues();
		_stateFeaturesBlocked = state.IsFeatureBlocked;
		_featuresBlocked = _stateFeaturesBlocked || _readOnlyMode;
		_canGenerate = state.CanGenerate;
		_preferredCreateNewTemplate = null;
		lblRemoteValue.Text = BuildOverallStatusText(state);
		_statusSummary.Text = _currentMappings.Count > 0
			? "已识别 " + _currentMappings.Count + " 个校准项"
			: "尚未识别到校准项";
		_matchChainLabel.Text = BuildMatchChainText(state);
		BindMappings();
		UpdateTemplateLibraryButtons();
		_isBinding = false;
	}

	public void BeginEditSavedTemplate(TaskPaneState state, SavedTemplateInfo template)
	{
		if (state == null || template == null) return;
		Bind(state);
		_isEditingSavedTemplate = true;
		_editingTemplateFingerprint = template.ExactFingerprint ?? string.Empty;
		_editingRemoteTemplateId = template.RemoteTemplateId ?? string.Empty;
		_editingTemplateName = template.TemplateName ?? string.Empty;
		_editingDirectoryMetadata = template.DirectoryMetadata;
		_preferredCreateNewTemplate = false;
		_hasUnsavedChanges = false;
		lblRemoteValue.Text = "正在编辑：" + (string.IsNullOrWhiteSpace(template.TemplateName) ? "未命名模板" : template.TemplateName);
		_statusSummary.Text = "已加载 " + _currentRules.Count + " 条已保存规则";
		UpdateTemplateLibraryButtons();
	}

	public bool IsEditingSavedTemplate => _isEditingSavedTemplate;
	public bool HasUnsavedChanges => _hasUnsavedChanges;

	public bool ConfirmCloseIfDirty()
	{
		if (!_hasUnsavedChanges) return true;
		var result = MessageBox.Show("当前模板有更改未保存，是否保存？", "模板编辑", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
		if (result == DialogResult.Cancel) return false;
		if (result == DialogResult.No) { _hasUnsavedChanges = false; return true; }
		return SaveEditedTemplate(false);
	}

	public IReadOnlyList<MeasurementRule> GetCurrentRules()
	{
		return (_featuresBlocked || !_canGenerate)
			? new List<MeasurementRule>()
			: TaskPaneModelCloner.CloneRules(_currentRules).Where(rule => rule.IsEnabled).ToList();
	}

	public void SetReadOnlyMode(bool readOnly)
	{
		_readOnlyMode = readOnly;
		_featuresBlocked = _stateFeaturesBlocked || _readOnlyMode;
		if (IsHandleCreated && _currentMappings.Count > 0)
		{
			BindMappings();
			UpdateTemplateLibraryButtons();
		}
	}

	public void SetPreferredCreateNewTemplate(bool? createNew)
	{
		_preferredCreateNewTemplate = createNew;
	}

	private void NotifyGenerationStateChanged()
	{
		if (!_isBinding) _hasUnsavedChanges = true;
		if (_featuresBlocked)
		{
			return;
		}

		try
		{
			Globals.ThisAddIn.UpdateCurrentGenerationRules(
				GetCurrentRules(),
				CloneGenerationConfiguration(_appliedGenerationConfiguration));
		}
		catch
		{
		}
	}

	private IReadOnlyList<MeasurementRule> GetSavableRules()
	{
		return _featuresBlocked
			? new List<MeasurementRule>()
			: TaskPaneModelCloner.CloneRules(_currentRules).Where(rule => rule.IsEnabled).ToList();
	}

	public ExcelCalibrationAddin.Core.Models.GenerationConfiguration GetAppliedGenerationConfiguration()
	{
		return CloneGenerationConfiguration(_appliedGenerationConfiguration);
	}

	public bool IsFeatureBlocked()
	{
		return _featuresBlocked;
	}

	public void ClearMappingSelection()
	{
		ClearFieldStatusSelection();
	}

	public void SetRecognitionProgress(string message, int percent, bool visible)
	{
		if (base.InvokeRequired)
		{
			BeginInvoke((MethodInvoker)delegate
			{
				SetRecognitionProgress(message, percent, visible);
			});
			return;
		}
		_progressLabel.Text = (string.IsNullOrWhiteSpace(message) ? "正在识别..." : message);
		_progressBar.Value = ((percent >= 0) ? ((percent > 100) ? 100 : percent) : 0);
		_progressLabel.Visible = visible;
		_progressBar.Visible = visible;
		_statusSummary.Visible = !visible;
	}

    }
}
