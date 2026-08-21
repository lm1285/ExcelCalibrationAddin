using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Windows.Forms;
using ExcelCalibrationAddin.Core.Models;

namespace ExcelCalibrationAddin.Vsto
{
    internal sealed class RandomGenerationConfigurationDialog : Form
    {
        private static readonly Color WindowColor = Color.FromArgb(245, 245, 247);
        private static readonly Color SurfaceColor = Color.White;
        private static readonly Color FieldColor = Color.FromArgb(249, 250, 251);
        private static readonly Color LineColor = Color.FromArgb(226, 226, 230);
        private static readonly Color TextColor = Color.FromArgb(29, 29, 31);
        private static readonly Color MutedColor = Color.FromArgb(105, 105, 112);
        private static readonly Color AccentColor = Color.FromArgb(0, 137, 143);
        private static readonly Color AccentHoverColor = Color.FromArgb(0, 108, 113);
        private const float FieldLabelWidth = 72F;
        private const float FieldInputWidth = 180F;

        private readonly Button _sameDirectionEnableButton = new RoundedButton();
        private readonly Button _independentDeviationEnableButton = new RoundedButton();
        private readonly ModernNumericInput _unifiedMinInput = new ModernNumericInput();
        private readonly ModernNumericInput _unifiedMaxInput = new ModernNumericInput();
        private readonly ModernNumericInput _positiveMinInput = new ModernNumericInput();
        private readonly ModernNumericInput _positiveMaxInput = new ModernNumericInput();
        private readonly ModernNumericInput _negativeMinInput = new ModernNumericInput();
        private readonly ModernNumericInput _negativeMaxInput = new ModernNumericInput();
        private readonly ModernNumericInput _absoluteMinInput = new ModernNumericInput();
        private readonly ModernNumericInput _absoluteMaxInput = new ModernNumericInput();
        private readonly ModernNumericInput _minimumRequirementMinInput = new ModernNumericInput();
        private readonly ModernNumericInput _minimumRequirementMaxInput = new ModernNumericInput();
        private readonly ModernNumericInput _measurementGroupMinimumFluctuationInput = new ModernNumericInput();
        private readonly ModernNumericInput _measurementGroupMaximumFluctuationInput = new ModernNumericInput();
        private readonly ModernNumericInput _resultGroupMinimumFluctuationInput = new ModernNumericInput();
        private readonly ModernNumericInput _resultGroupMaximumFluctuationInput = new ModernNumericInput();
        private readonly ModernNumericInput _responseTimeThresholdInput = new ModernNumericInput();
        private readonly ModernNumericInput _responseTimeBelowThresholdDifferenceInput = new ModernNumericInput();
        private readonly ModernNumericInput _responseTimeAboveThresholdDifferenceInput = new ModernNumericInput();
        private readonly Panel _unifiedDeviationPanel = new Panel();
        private readonly Panel _independentDeviationPanel = new Panel();
        private readonly Panel _tabNavigation = new Panel();
        private readonly Panel _tabIndicator = new Panel();
        private readonly Timer _tabAnimationTimer = new Timer();
        private readonly ToolTip _toolTip = new ToolTip();
        private Button _deviationTabButton;
        private Button _fluctuationTabButton;
        private Button _responseTabButton;
        private Panel _deviationPage;
        private Panel _fluctuationPage;
        private Panel _responsePage;
        private readonly TextBox _shortcutKeyInput = new TextBox();
        private readonly Button _clearShortcutButton = new RoundedButton();
        private int _selectedTabIndex;
        private int _indicatorTargetLeft;

        public RandomGenerationConfigurationDialog(GenerationConfiguration configuration)
        {
            Configuration = Clone(configuration);
            Text = "测量值生成配置";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(780, 820);
            BackColor = WindowColor;
            Font = CreateUiFont(9F, FontStyle.Regular);
            DoubleBuffered = true;

            _tabAnimationTimer.Interval = 15;
            _tabAnimationTimer.Tick += (_, __) => AnimateTabIndicator();
            _toolTip.AutoPopDelay = 12000;
            _toolTip.InitialDelay = 250;
            _toolTip.ReshowDelay = 100;

            BuildLayout();
            Bind(Configuration);
        }

        public GenerationConfiguration Configuration { get; private set; }

        private float UiScale
        {
            get
            {
                return Math.Max(1F, Math.Min(1.2F, ClientSize.Width / 640F));
            }
        }

        private Font CreateUiFont(float size, FontStyle style)
        {
            return new Font("Microsoft YaHei UI", size * UiScale, style);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _tabAnimationTimer.Dispose();
                _toolTip.Dispose();
            }

            base.Dispose(disposing);
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = WindowColor,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 78F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 66F));
            Controls.Add(root);

            root.Controls.Add(CreateTitleBar(), 0, 0);
            root.Controls.Add(CreateTabNavigation(), 0, 1);
            root.Controls.Add(CreateContentHost(), 0, 2);
            root.Controls.Add(CreateFooter(), 0, 3);
        }

        private Control CreateTitleBar()
        {
            var titleBar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = SurfaceColor,
                Margin = Padding.Empty,
                Padding = new Padding(36, 22, 36, 15)
            };
            titleBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            titleBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var title = CreateTextLabel("测量值生成配置", 18F, FontStyle.Bold, TextColor);
            title.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
            titleBar.Controls.Add(title, 0, 0);

            var context = CreateTextLabel("随机数配置", 9F, FontStyle.Regular, MutedColor);
            context.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            titleBar.Controls.Add(context, 1, 0);
            return titleBar;
        }

        private Control CreateTabNavigation()
        {
            _tabNavigation.BackColor = SurfaceColor;
            _tabNavigation.Dock = DockStyle.Fill;
            _tabNavigation.Margin = Padding.Empty;
            _tabNavigation.Padding = new Padding(36, 0, 36, 0);

            var tabGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SurfaceColor
            };
            tabGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333F));
            tabGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333F));
            tabGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333F));

            _deviationTabButton = CreateTabButton("允差控制");
            _fluctuationTabButton = CreateTabButton("波动控制");
            _responseTabButton = CreateTabButton("响应时间");
            _deviationTabButton.Click += (_, __) => SelectTab(0);
            _fluctuationTabButton.Click += (_, __) => SelectTab(1);
            _responseTabButton.Click += (_, __) => SelectTab(2);
            tabGrid.Controls.Add(_deviationTabButton, 0, 0);
            tabGrid.Controls.Add(_fluctuationTabButton, 1, 0);
            tabGrid.Controls.Add(_responseTabButton, 2, 0);
            _tabNavigation.Controls.Add(tabGrid);

            _tabIndicator.BackColor = AccentColor;
            _tabIndicator.Height = 3;
            _tabIndicator.Width = 44;
            _tabIndicator.Anchor = AnchorStyles.Bottom;
            _tabNavigation.Controls.Add(_tabIndicator);
            _tabNavigation.Resize += (_, __) => PositionTabIndicator(false);
            _tabIndicator.BringToFront();
            return _tabNavigation;
        }

        private Control CreateContentHost()
        {
            var host = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceColor,
                Margin = Padding.Empty,
                Padding = new Padding(36, 28, 36, 20),
                AutoScroll = true
            };

            _deviationPage = CreatePage();
            _fluctuationPage = CreatePage();
            _responsePage = CreatePage();
            _deviationPage.Controls.Add(CreateDeviationContent());
            _fluctuationPage.Controls.Add(CreateFluctuationContent());
            _responsePage.Controls.Add(CreateResponseContent());
            _fluctuationPage.Visible = false;
            _responsePage.Visible = false;
            host.Controls.Add(_responsePage);
            host.Controls.Add(_fluctuationPage);
            host.Controls.Add(_deviationPage);
            return host;
        }

        private static Panel CreatePage()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SurfaceColor,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoScroll = true
            };
        }

        private Control CreateDeviationContent()
        {
            var grid = CreatePageGrid();
            AddPageHeading(grid, "偏差方向与允差区间");
            AddDivider(grid);
            grid.Controls.Add(CreateShortcutSection());

            ConfigureEnableButton(_sameDirectionEnableButton);
            grid.Controls.Add(CreateSettingRow(
                "同一校准项保持同一偏差方向",
                _sameDirectionEnableButton,
                "同一校准项的随机值将保持正向或负向一致。\r\n适用于需要控制同一组数据偏差趋势的场景，关闭后每个测量值会独立取样。"));

            ConfigureEnableButton(_independentDeviationEnableButton);
            grid.Controls.Add(CreateSettingRow(
                "正负允许误差分别控制",
                _independentDeviationEnableButton,
                "为正、负偏差分别设置独立的占用比例区间。\r\n启用后，正偏差和负偏差使用各自的下限、上限；未启用时共用同一对区间。"));

            _unifiedDeviationPanel.Controls.Clear();
            _unifiedDeviationPanel.AutoSize = true;
            _unifiedDeviationPanel.Dock = DockStyle.Top;
            _unifiedDeviationPanel.Margin = Padding.Empty;
            _unifiedDeviationPanel.Padding = new Padding(0, 18, 0, 18);
            _unifiedDeviationPanel.Controls.Add(CreateRangeSection(
                "±允差区间",
                "控制随机值相对技术要求的偏差范围。",
                _unifiedMinInput,
                _unifiedMaxInput,
                0M,
                1M));
            grid.Controls.Add(_unifiedDeviationPanel);

            _independentDeviationPanel.Controls.Clear();
            _independentDeviationPanel.AutoSize = true;
            _independentDeviationPanel.Dock = DockStyle.Top;
            _independentDeviationPanel.Margin = Padding.Empty;
            _independentDeviationPanel.Padding = new Padding(0, 18, 0, 18);
            _independentDeviationPanel.Controls.Add(CreateIndependentRangeSection());
            grid.Controls.Add(_independentDeviationPanel);

            grid.Controls.Add(CreateRangeSection(
                "≤/＜允差控制",
                "用于技术要求为 ≤ 或 ＜ 的项目。",
                _absoluteMinInput,
                _absoluteMaxInput,
                0M,
                1M));
            grid.Controls.Add(CreateRangeSection(
                "＞/≥允差控制",
                "用于技术要求为 ＞ 或 ≥ 的项目。",
                _minimumRequirementMinInput,
                _minimumRequirementMaxInput,
                1M,
                10M));
            return grid;
        }

        private Control CreateShortcutSection()
        {
            var section = CreateSectionContainer();
            var row = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 3,
                RowCount = 2,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SurfaceColor
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));

            var title = CreateTextLabel("生成随机数快捷键", 9F, FontStyle.Bold, TextColor);
            title.AutoSize = true;
            title.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            title.Margin = Padding.Empty;
            row.Controls.Add(title, 0, 0);

            _shortcutKeyInput.ReadOnly = true;
            _shortcutKeyInput.ShortcutsEnabled = false;
            _shortcutKeyInput.TextAlign = HorizontalAlignment.Center;
            _shortcutKeyInput.BorderStyle = BorderStyle.FixedSingle;
            _shortcutKeyInput.BackColor = FieldColor;
            _shortcutKeyInput.ForeColor = TextColor;
            _shortcutKeyInput.Font = CreateUiFont(9F, FontStyle.Bold);
            _shortcutKeyInput.Height = 30;
            _shortcutKeyInput.Dock = DockStyle.Fill;
            _shortcutKeyInput.Margin = new Padding(0, 0, 10, 0);
            _shortcutKeyInput.KeyDown += ShortcutKeyInput_KeyDown;
            row.Controls.Add(_shortcutKeyInput, 1, 0);

            _clearShortcutButton.Text = "清空";
            _clearShortcutButton.Width = 60;
            _clearShortcutButton.Height = 30;
            _clearShortcutButton.Margin = Padding.Empty;
            _clearShortcutButton.Font = CreateUiFont(8.5F, FontStyle.Regular);
            _clearShortcutButton.FlatStyle = FlatStyle.Flat;
            _clearShortcutButton.FlatAppearance.BorderSize = 1;
            _clearShortcutButton.FlatAppearance.BorderColor = LineColor;
            _clearShortcutButton.BackColor = SurfaceColor;
            _clearShortcutButton.ForeColor = TextColor;
            _clearShortcutButton.Click += (_, __) => _shortcutKeyInput.Clear();
            row.Controls.Add(_clearShortcutButton, 2, 0);

            var description = CreateTextLabel(
                "按下单个功能键，或 Ctrl/Alt/Shift + 一个按键进行设置；留空表示禁用。默认 F6。",
                8.5F,
                FontStyle.Regular,
                MutedColor);
            description.AutoSize = true;
            description.Margin = new Padding(0, 8, 0, 0);
            row.Controls.Add(description, 0, 1);
            row.SetColumnSpan(description, 3);

            section.Controls.Add(row);
            section.Controls.Add(CreateSectionHeader(
                "快捷键",
                "快捷键只在 Excel 主窗口生效；组合键按下指定修饰键后触发。",
                "单键或组合键"));
            return section;
        }

        private void ShortcutKeyInput_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;
            if (IsModifierKey(e.KeyCode))
            {
                return;
            }

            if (e.Modifiers == Keys.None && !IsFunctionKey(e.KeyCode))
            {
                return;
            }

            var parts = new System.Collections.Generic.List<string>();
            if ((e.Modifiers & Keys.Control) != Keys.None) parts.Add("Ctrl");
            if ((e.Modifiers & Keys.Alt) != Keys.None) parts.Add("Alt");
            if ((e.Modifiers & Keys.Shift) != Keys.None) parts.Add("Shift");
            parts.Add(e.KeyCode.ToString().ToUpperInvariant());
            _shortcutKeyInput.Text = string.Join("+", parts);
        }

        private static bool IsModifierKey(Keys key)
        {
            return key == Keys.Control || key == Keys.Menu || key == Keys.ShiftKey;
        }

        private static bool IsFunctionKey(Keys key)
        {
            return key >= Keys.F1 && key <= Keys.F12;
        }

        private Control CreateFluctuationContent()
        {
            var grid = CreatePageGrid();
            AddPageHeading(grid, "标准值与结果波动");
            AddDivider(grid);
            grid.Controls.Add(CreateRangeSection(
                "单组标准值内测量值最大波动",
                "控制同一组标准值内测量值的最大波动。",
                _measurementGroupMinimumFluctuationInput,
                _measurementGroupMaximumFluctuationInput,
                0.01M,
                1M));
            grid.Controls.Add(CreateRangeSection(
                "多组标准值误差的最大波动",
                "控制多组标准值之间误差的最大波动。",
                _resultGroupMinimumFluctuationInput,
                _resultGroupMaximumFluctuationInput,
                0M,
                1M));
            return grid;
        }

        private Control CreateResponseContent()
        {
            var grid = CreatePageGrid();
            AddPageHeading(grid, "响应时间差值控制");
            AddDivider(grid);
            grid.Controls.Add(CreateSingleRangeSection(
                "响应时间阈值",
                "阈值 ≤",
                _responseTimeThresholdInput,
                0.01M,
                100000M,
                "响应时间达到该阈值时，使用阈值内的差值控制。"));
            grid.Controls.Add(CreateSingleRangeSection(
                "阈值内最大差值",
                "差值 ≤",
                _responseTimeBelowThresholdDifferenceInput,
                0.01M,
                100000M,
                "响应时间不超过阈值时允许的最大差值。"));
            grid.Controls.Add(CreateSingleRangeSection(
                "阈值外最大差值",
                "差值 ≤",
                _responseTimeAboveThresholdDifferenceInput,
                0.01M,
                100000M,
                "响应时间超过阈值时允许的最大差值。"));
            return grid;
        }

        private TableLayoutPanel CreatePageGrid()
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 0,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SurfaceColor
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            return grid;
        }

        private void AddPageHeading(TableLayoutPanel grid, string text)
        {
            var heading = CreateTextLabel(text, 15F, FontStyle.Bold, TextColor);
            heading.AutoSize = true;
            heading.Margin = Padding.Empty;
            grid.Controls.Add(heading);
        }

        private static void AddDivider(TableLayoutPanel grid)
        {
            var divider = new Panel
            {
                Height = 17,
                Dock = DockStyle.Top,
                BackColor = SurfaceColor,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            divider.Controls.Add(new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 1,
                BackColor = LineColor,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            });
            grid.Controls.Add(divider);
        }

        private TableLayoutPanel CreateSettingRow(string title, Button enableButton, string info)
        {
            var row = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = new Padding(0, 16, 0, 16),
                BackColor = SurfaceColor
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 70F));
            var titleBlock = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SurfaceColor
            };
            titleBlock.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            titleBlock.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30F));
            var titleLabel = CreateTextLabel(title, 9F, FontStyle.Bold, TextColor);
            titleLabel.AutoSize = true;
            titleLabel.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            titleLabel.Margin = Padding.Empty;
            titleBlock.Controls.Add(titleLabel, 0, 0);
            var infoButton = CreateInfoButton(info);
            infoButton.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            titleBlock.Controls.Add(infoButton, 1, 0);
            row.Controls.Add(titleBlock, 0, 0);
            enableButton.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            enableButton.Margin = Padding.Empty;
            row.Controls.Add(enableButton, 1, 0);
            return row;
        }

        private Control CreateRangeSection(
            string title,
            string info,
            ModernNumericInput minimumInput,
            ModernNumericInput maximumInput,
            decimal minimum,
            decimal maximum)
        {
            ConfigureNumber(minimumInput, minimum, maximum);
            ConfigureNumber(maximumInput, minimum, maximum);
            var section = CreateSectionContainer();
            section.Controls.Add(CreateRangeGrid(minimumInput, maximumInput));
            section.Controls.Add(CreateSectionHeader(title, info, "技术要求 × 系数"));
            return section;
        }

        private Control CreateIndependentRangeSection()
        {
            ConfigureNumber(_positiveMinInput, 0M, 1M);
            ConfigureNumber(_positiveMaxInput, 0M, 1M);
            ConfigureNumber(_negativeMinInput, 0M, 1M);
            ConfigureNumber(_negativeMaxInput, 0M, 1M);
            var section = CreateSectionContainer();
            var splitGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SurfaceColor
            };
            splitGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            splitGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            splitGrid.Controls.Add(CreateIndependentRangeBlock("正偏差", _positiveMinInput, _positiveMaxInput), 0, 0);
            splitGrid.Controls.Add(CreateIndependentRangeBlock("负偏差", _negativeMinInput, _negativeMaxInput), 1, 0);
            section.Controls.Add(splitGrid);
            section.Controls.Add(CreateSectionHeader(
                "±允差区间",
                "启用后，正偏差和负偏差各自使用独立区间。",
                "技术要求 × 系数"));
            return section;
        }

        private Control CreateIndependentRangeBlock(string title, ModernNumericInput minimumInput, ModernNumericInput maximumInput)
        {
            var block = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 3,
                Margin = Padding.Empty,
                Padding = new Padding(0, 0, 16, 0),
                BackColor = SurfaceColor
            };
            block.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FieldLabelWidth));
            block.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FieldInputWidth));
            block.Controls.Add(CreateTextLabel(title, 9F, FontStyle.Bold, TextColor), 0, 0);
            block.SetColumnSpan(block.GetControlFromPosition(0, 0), 2);
            block.Controls.Add(CreateFieldLabel("下限"), 0, 1);
            block.Controls.Add(CreateNumericField(minimumInput), 1, 1);
            block.Controls.Add(CreateFieldLabel("上限"), 0, 2);
            block.Controls.Add(CreateNumericField(maximumInput), 1, 2);
            return block;
        }

        private Control CreateSingleRangeSection(
            string title,
            string label,
            ModernNumericInput input,
            decimal minimum,
            decimal maximum,
            string info)
        {
            ConfigureNumber(input, minimum, maximum);
            var section = CreateSectionContainer();
            var fields = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SurfaceColor
            };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FieldLabelWidth));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FieldInputWidth));
            fields.Controls.Add(CreateFieldLabel(label), 0, 0);
            fields.Controls.Add(CreateNumericField(input), 1, 0);
            section.Controls.Add(fields);
            section.Controls.Add(CreateSectionHeader(title, info, "秒"));
            return section;
        }

        private static Panel CreateSectionContainer()
        {
            return new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                Margin = Padding.Empty,
                Padding = new Padding(0, 18, 0, 18),
                BackColor = SurfaceColor
            };
        }

        private Control CreateSectionHeader(string title, string info, string unit)
        {
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 3,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = new Padding(0, 0, 0, 14),
                BackColor = SurfaceColor
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30F));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            var titleLabel = CreateTextLabel(title, 9F, FontStyle.Bold, TextColor);
            titleLabel.AutoSize = true;
            titleLabel.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            header.Controls.Add(titleLabel, 0, 0);
            var tooltip = string.IsNullOrWhiteSpace(unit)
                ? info
                : info + "\r\n单位：" + unit;
            var infoButton = CreateInfoButton(tooltip);
            infoButton.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            header.Controls.Add(infoButton, 1, 0);
            return header;
        }

        private TableLayoutPanel CreateRangeGrid(ModernNumericInput minimumInput, ModernNumericInput maximumInput)
        {
            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SurfaceColor
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            grid.Controls.Add(CreateRangeField("下限", minimumInput), 0, 0);
            grid.Controls.Add(CreateRangeField("上限", maximumInput), 1, 0);
            return grid;
        }

        private Control CreateRangeField(string label, ModernNumericInput input)
        {
            var field = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 2,
                RowCount = 1,
                Margin = Padding.Empty,
                Padding = new Padding(0, 0, 16, 0),
                BackColor = SurfaceColor
            };
            field.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FieldLabelWidth));
            field.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FieldInputWidth));
            field.Controls.Add(CreateFieldLabel(label), 0, 0);
            field.Controls.Add(CreateNumericField(input), 1, 0);
            return field;
        }

        private Label CreateFieldLabel(string text)
        {
            var label = CreateTextLabel(text, 8.5F, FontStyle.Regular, MutedColor);
            label.AutoSize = true;
            label.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            label.Margin = new Padding(0, 1, 0, 0);
            return label;
        }

        private Button CreateInfoButton(string text)
        {
            var button = new InfoButton
            {
                Text = string.Empty,
                Width = 22,
                Height = 22,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = Padding.Empty,
                TabStop = true,
                AccessibleName = "说明",
                Font = CreateUiFont(9F, FontStyle.Bold)
            };
            _toolTip.SetToolTip(button, text);
            return button;
        }

        private Label CreateTextLabel(string text, float size, FontStyle style, Color color)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = color,
                Font = CreateUiFont(size, style),
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
        }

        private Button CreateTabButton(string text)
        {
            return new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                FlatAppearance = { BorderSize = 0 },
                BackColor = SurfaceColor,
                ForeColor = MutedColor,
                Font = CreateUiFont(9F, FontStyle.Regular),
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                Cursor = Cursors.Hand,
                TabStop = true
            };
        }

        private void ConfigureEnableButton(Button button)
        {
            button.Text = "启用";
            button.Width = 58;
            button.Height = 30;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.Font = CreateUiFont(8.5F, FontStyle.Bold);
            button.Cursor = Cursors.Hand;
            var roundedButton = button as RoundedButton;
            if (roundedButton != null)
            {
                roundedButton.BorderColor = LineColor;
                roundedButton.HoverBackColor = Color.FromArgb(248, 248, 250);
            }
            SetEnableButtonState(button, false);
        }

        private static void SetEnableButtonState(Button button, bool enabled)
        {
            button.Tag = enabled;
            button.BackColor = enabled ? AccentColor : SurfaceColor;
            button.ForeColor = enabled ? Color.White : MutedColor;
            button.FlatAppearance.BorderColor = enabled ? AccentColor : LineColor;
            button.FlatAppearance.MouseOverBackColor = enabled ? AccentHoverColor : Color.FromArgb(248, 248, 250);
            var roundedButton = button as RoundedButton;
            if (roundedButton != null)
            {
                roundedButton.BorderColor = enabled ? AccentColor : LineColor;
                roundedButton.HoverBackColor = enabled ? AccentHoverColor : Color.FromArgb(248, 248, 250);
                roundedButton.Invalidate();
            }
        }

        private void ConfigureNumber(ModernNumericInput input, decimal minimum, decimal maximum)
        {
            input.Minimum = minimum;
            input.Maximum = maximum;
            input.DecimalPlaces = 2;
            input.Increment = 0.01M;
            input.BackColor = FieldColor;
            input.ForeColor = TextColor;
            input.Font = new Font("Consolas", 9F, FontStyle.Regular);
            input.Height = 30;
            input.Width = 118;
            input.Margin = Padding.Empty;
        }

        private Control CreateNumericField(ModernNumericInput input)
        {
            input.Font = new Font("Consolas", 10F * UiScale, FontStyle.Regular);
            input.Height = (int)Math.Round(36F * UiScale);
            input.Width = (int)Math.Round(FieldInputWidth * UiScale);
            input.Dock = DockStyle.Left;
            input.Anchor = AnchorStyles.Left | AnchorStyles.Top;
            input.Margin = Padding.Empty;
            return input;
        }

        private Control CreateFooter()
        {
            var footer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = SurfaceColor,
                Margin = Padding.Empty,
                Padding = new Padding(36, 15, 36, 18)
            };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82F));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82F));
            var cancelButton = CreateActionButton("取消", false);
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            var okButton = CreateActionButton("保存", true);
            okButton.Click += (_, __) => Accept();
            okButton.Anchor = AnchorStyles.Right | AnchorStyles.Top;
            footer.Controls.Add(new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty }, 0, 0);
            footer.Controls.Add(cancelButton, 1, 0);
            footer.Controls.Add(okButton, 2, 0);
            AcceptButton = okButton;
            CancelButton = cancelButton;
            return footer;
        }

        private Button CreateActionButton(string text, bool primary)
        {
            var button = new RoundedButton
            {
                Text = text,
                Width = 82,
                Height = 34,
                BackColor = primary ? AccentColor : SurfaceColor,
                ForeColor = primary ? Color.White : TextColor,
                BorderColor = primary ? AccentColor : LineColor,
                HoverBackColor = primary ? AccentHoverColor : Color.FromArgb(248, 248, 250),
                Font = CreateUiFont(9F, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                Margin = Padding.Empty,
                Cursor = Cursors.Hand
            };
            return button;
        }

        private void SelectTab(int index)
        {
            _selectedTabIndex = index;
            _deviationPage.Visible = index == 0;
            _fluctuationPage.Visible = index == 1;
            _responsePage.Visible = index == 2;
            _deviationTabButton.ForeColor = index == 0 ? TextColor : MutedColor;
            _fluctuationTabButton.ForeColor = index == 1 ? TextColor : MutedColor;
            _responseTabButton.ForeColor = index == 2 ? TextColor : MutedColor;
            PositionTabIndicator(true);
        }

        private void PositionTabIndicator(bool animate)
        {
            var availableWidth = Math.Max(0, _tabNavigation.ClientSize.Width - _tabNavigation.Padding.Left - _tabNavigation.Padding.Right);
            var cellWidth = availableWidth / 3;
            _indicatorTargetLeft = _tabNavigation.Padding.Left + cellWidth * _selectedTabIndex + Math.Max(0, (cellWidth - _tabIndicator.Width) / 2);
            if (!animate)
            {
                _tabIndicator.Left = _indicatorTargetLeft;
                _tabIndicator.Top = _tabNavigation.ClientSize.Height - _tabIndicator.Height;
                return;
            }

            _tabAnimationTimer.Start();
        }

        private void AnimateTabIndicator()
        {
            var distance = _indicatorTargetLeft - _tabIndicator.Left;
            if (Math.Abs(distance) <= 1)
            {
                _tabIndicator.Left = _indicatorTargetLeft;
                _tabAnimationTimer.Stop();
                return;
            }

            _tabIndicator.Left += Math.Sign(distance) * Math.Max(1, Math.Abs(distance) / 4);
            _tabIndicator.Top = _tabNavigation.ClientSize.Height - _tabIndicator.Height;
        }

        private void Bind(GenerationConfiguration configuration)
        {
            SetEnableButtonState(_sameDirectionEnableButton, configuration.UseSameDeviationDirection);
            SetEnableButtonState(_independentDeviationEnableButton, configuration.UseIndependentDeviationControl);
            _sameDirectionEnableButton.Click += (_, __) => SetEnableButtonState(_sameDirectionEnableButton, !IsEnabled(_sameDirectionEnableButton));
            _independentDeviationEnableButton.Click += (_, __) =>
            {
                SetEnableButtonState(_independentDeviationEnableButton, !IsEnabled(_independentDeviationEnableButton));
                UpdateDeviationPanels();
            };
            _unifiedMinInput.Value = ClampCoefficient(configuration.UnifiedErrorMinimumCoefficient, _unifiedMinInput);
            _unifiedMaxInput.Value = ClampCoefficient(configuration.UnifiedErrorMaximumCoefficient, _unifiedMaxInput);
            _positiveMinInput.Value = ClampCoefficient(configuration.PositiveErrorMinimumCoefficient, _positiveMinInput);
            _positiveMaxInput.Value = ClampCoefficient(configuration.PositiveErrorMaximumCoefficient, _positiveMaxInput);
            _negativeMinInput.Value = ClampCoefficient(configuration.NegativeErrorMinimumCoefficient, _negativeMinInput);
            _negativeMaxInput.Value = ClampCoefficient(configuration.NegativeErrorMaximumCoefficient, _negativeMaxInput);
            _absoluteMinInput.Value = ClampCoefficient(configuration.AbsoluteErrorMinimumCoefficient, _absoluteMinInput);
            _absoluteMaxInput.Value = ClampCoefficient(configuration.AbsoluteErrorMaximumCoefficient, _absoluteMaxInput);
            _minimumRequirementMinInput.Value = ClampCoefficient(configuration.MinimumRequirementMinimumCoefficient, _minimumRequirementMinInput);
            _minimumRequirementMaxInput.Value = ClampCoefficient(configuration.MinimumRequirementMaximumCoefficient, _minimumRequirementMaxInput);
            _measurementGroupMinimumFluctuationInput.Value = ClampCoefficient(configuration.MeasurementGroupMinimumFluctuationCoefficient, _measurementGroupMinimumFluctuationInput);
            _measurementGroupMaximumFluctuationInput.Value = ClampCoefficient(configuration.MeasurementGroupMaximumFluctuationCoefficient, _measurementGroupMaximumFluctuationInput);
            _resultGroupMinimumFluctuationInput.Value = ClampCoefficient(configuration.ResultGroupMinimumFluctuationCoefficient, _resultGroupMinimumFluctuationInput);
            _resultGroupMaximumFluctuationInput.Value = ClampCoefficient(configuration.ResultGroupMaximumFluctuationCoefficient, _resultGroupMaximumFluctuationInput);
            _responseTimeThresholdInput.Value = ClampCoefficient(configuration.ResponseTimeThresholdSeconds, _responseTimeThresholdInput);
            _responseTimeBelowThresholdDifferenceInput.Value = ClampCoefficient(configuration.ResponseTimeBelowThresholdMaximumDifferenceSeconds, _responseTimeBelowThresholdDifferenceInput);
            _responseTimeAboveThresholdDifferenceInput.Value = ClampCoefficient(configuration.ResponseTimeAboveThresholdMaximumDifferenceSeconds, _responseTimeAboveThresholdDifferenceInput);
            _shortcutKeyInput.Text = configuration.GenerateShortcutKey ?? string.Empty;
            UpdateDeviationPanels();
            SelectTab(0);
        }

        private void UpdateDeviationPanels()
        {
            var useIndependentControl = IsEnabled(_independentDeviationEnableButton);
            _independentDeviationPanel.Visible = useIndependentControl;
            _unifiedDeviationPanel.Visible = !useIndependentControl;
            _unifiedDeviationPanel.Parent?.PerformLayout();
            _independentDeviationPanel.Parent?.PerformLayout();
        }

        private static bool IsEnabled(Button button)
        {
            return button.Tag is bool && (bool)button.Tag;
        }

        private void Accept()
        {
            var independent = IsEnabled(_independentDeviationEnableButton);
            if ((!independent && !ValidateCoefficientRange(_unifiedMinInput, _unifiedMaxInput, "允许误差占用比例")) ||
                (independent && (!ValidateCoefficientRange(_positiveMinInput, _positiveMaxInput, "正偏差占用比例") ||
                                 !ValidateCoefficientRange(_negativeMinInput, _negativeMaxInput, "负偏差占用比例"))) ||
                !ValidateCoefficientRange(_absoluteMinInput, _absoluteMaxInput, "≤/＜允差控制") ||
                !ValidateCoefficientRange(_minimumRequirementMinInput, _minimumRequirementMaxInput, "＞/≥允差控制") ||
                !ValidateCoefficientRange(_measurementGroupMinimumFluctuationInput, _measurementGroupMaximumFluctuationInput, "单组标准值内测量值最大波动") ||
                !ValidateCoefficientRange(_resultGroupMinimumFluctuationInput, _resultGroupMaximumFluctuationInput, "多组标准值误差的最大波动"))
            {
                return;
            }

            Configuration = new GenerationConfiguration
            {
                GenerateShortcutKey = (_shortcutKeyInput.Text ?? string.Empty).Trim().ToUpperInvariant(),
                DefaultDistribution = "Normal",
                ResultCalculationMethod = "FormulaBackCalculation",
                StandardValueReference = "RecognizedStandardValueRange",
                UseSameDeviationDirection = IsEnabled(_sameDirectionEnableButton),
                UseIndependentDeviationControl = independent,
                UnifiedErrorMinimumCoefficient = Convert.ToDouble(_unifiedMinInput.Value),
                UnifiedErrorMaximumCoefficient = Convert.ToDouble(_unifiedMaxInput.Value),
                PositiveErrorMinimumCoefficient = Convert.ToDouble(_positiveMinInput.Value),
                PositiveErrorMaximumCoefficient = Convert.ToDouble(_positiveMaxInput.Value),
                NegativeErrorMinimumCoefficient = Convert.ToDouble(_negativeMinInput.Value),
                NegativeErrorMaximumCoefficient = Convert.ToDouble(_negativeMaxInput.Value),
                AbsoluteErrorMinimumCoefficient = Convert.ToDouble(_absoluteMinInput.Value),
                AbsoluteErrorMaximumCoefficient = Convert.ToDouble(_absoluteMaxInput.Value),
                MinimumRequirementMinimumCoefficient = Convert.ToDouble(_minimumRequirementMinInput.Value),
                MinimumRequirementMaximumCoefficient = Convert.ToDouble(_minimumRequirementMaxInput.Value),
                UseDecimalPlacesForResolution = true,
                MeasurementGroupMinimumFluctuationCoefficient = Convert.ToDouble(_measurementGroupMinimumFluctuationInput.Value),
                MeasurementGroupMaximumFluctuationCoefficient = Convert.ToDouble(_measurementGroupMaximumFluctuationInput.Value),
                ResultGroupMinimumFluctuationCoefficient = Convert.ToDouble(_resultGroupMinimumFluctuationInput.Value),
                ResultGroupMaximumFluctuationCoefficient = Convert.ToDouble(_resultGroupMaximumFluctuationInput.Value),
                ResponseTimeThresholdSeconds = Convert.ToDouble(_responseTimeThresholdInput.Value),
                ResponseTimeBelowThresholdMaximumDifferenceSeconds = Convert.ToDouble(_responseTimeBelowThresholdDifferenceInput.Value),
                ResponseTimeAboveThresholdMaximumDifferenceSeconds = Convert.ToDouble(_responseTimeAboveThresholdDifferenceInput.Value)
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        private bool ValidateCoefficientRange(ModernNumericInput minInput, ModernNumericInput maxInput, string label)
        {
            if (minInput.Value <= maxInput.Value)
            {
                return true;
            }

            MessageBox.Show($"{label}的下限不能大于上限。", "测量值生成配置", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            minInput.Focus();
            return false;
        }

        private static GenerationConfiguration Clone(GenerationConfiguration configuration)
        {
            if (configuration == null)
            {
                return new GenerationConfiguration();
            }

            return new GenerationConfiguration
            {
                GenerateShortcutKey = configuration.GenerateShortcutKey,
                DefaultDistribution = configuration.DefaultDistribution,
                ResultCalculationMethod = configuration.ResultCalculationMethod,
                StandardValueReference = configuration.StandardValueReference,
                UseSameDeviationDirection = configuration.UseSameDeviationDirection,
                UseIndependentDeviationControl = configuration.UseIndependentDeviationControl,
                UnifiedErrorMinimumCoefficient = configuration.UnifiedErrorMinimumCoefficient,
                UnifiedErrorMaximumCoefficient = configuration.UnifiedErrorMaximumCoefficient,
                PositiveErrorMinimumCoefficient = configuration.PositiveErrorMinimumCoefficient,
                PositiveErrorMaximumCoefficient = configuration.PositiveErrorMaximumCoefficient,
                NegativeErrorMinimumCoefficient = configuration.NegativeErrorMinimumCoefficient,
                NegativeErrorMaximumCoefficient = configuration.NegativeErrorMaximumCoefficient,
                AbsoluteErrorMinimumCoefficient = configuration.AbsoluteErrorMinimumCoefficient,
                AbsoluteErrorMaximumCoefficient = configuration.AbsoluteErrorMaximumCoefficient,
                MinimumRequirementMinimumCoefficient = configuration.MinimumRequirementMinimumCoefficient,
                MinimumRequirementMaximumCoefficient = configuration.MinimumRequirementMaximumCoefficient,
                UseDecimalPlacesForResolution = configuration.UseDecimalPlacesForResolution,
                MeasurementGroupMinimumFluctuationCoefficient = configuration.MeasurementGroupMinimumFluctuationCoefficient,
                MeasurementGroupMaximumFluctuationCoefficient = configuration.MeasurementGroupMaximumFluctuationCoefficient,
                ResultGroupMinimumFluctuationCoefficient = configuration.ResultGroupMinimumFluctuationCoefficient,
                ResultGroupMaximumFluctuationCoefficient = configuration.ResultGroupMaximumFluctuationCoefficient,
                ResponseTimeThresholdSeconds = configuration.ResponseTimeThresholdSeconds,
                ResponseTimeBelowThresholdMaximumDifferenceSeconds = configuration.ResponseTimeBelowThresholdMaximumDifferenceSeconds,
                ResponseTimeAboveThresholdMaximumDifferenceSeconds = configuration.ResponseTimeAboveThresholdMaximumDifferenceSeconds
            };
        }

        private static decimal ClampCoefficient(double value, ModernNumericInput input)
        {
            return Math.Max(input.Minimum, Math.Min(input.Maximum, Convert.ToDecimal(value)));
        }

        private sealed class RoundedButton : Button
        {
            public Color BorderColor { get; set; }
            public Color HoverBackColor { get; set; }
            private bool _hovered;

            protected override void OnMouseEnter(EventArgs e)
            {
                base.OnMouseEnter(e);
                _hovered = true;
                Invalidate();
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                _hovered = false;
                Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
                using (var path = CreatePath(bounds, 6))
                using (var brush = new SolidBrush(_hovered ? HoverBackColor : BackColor))
                using (var pen = new Pen(BorderColor))
                {
                    e.Graphics.FillPath(brush, path);
                    e.Graphics.DrawPath(pen, path);
                }

                TextRenderer.DrawText(e.Graphics, Text, Font, bounds, ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            }
        }

        private sealed class ModernNumericInput : UserControl
        {
            private readonly TextBox _textBox = new TextBox();
            private decimal _minimum;
            private decimal _maximum = 100M;
            private decimal _value;
            private int _decimalPlaces = 2;
            private bool _focused;

            public ModernNumericInput()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
                BackColor = FieldColor;
                ForeColor = TextColor;
                Padding = new Padding(9, 2, 9, 2);
                Margin = Padding.Empty;
                Height = 36;
                MinimumSize = new Size(100, 32);
                _textBox.BorderStyle = BorderStyle.None;
                _textBox.Dock = DockStyle.Fill;
                _textBox.BackColor = FieldColor;
                _textBox.ForeColor = TextColor;
                _textBox.TextAlign = HorizontalAlignment.Right;
                _textBox.Margin = Padding.Empty;
                _textBox.ShortcutsEnabled = true;
                _textBox.KeyPress += TextBoxKeyPress;
                _textBox.Enter += (_, __) =>
                {
                    _focused = true;
                    Invalidate();
                };
                _textBox.Leave += (_, __) =>
                {
                    _focused = false;
                    NormalizeText();
                    Invalidate();
                };
                Controls.Add(_textBox);
                Value = 0M;
            }

            public decimal Minimum
            {
                get { return _minimum; }
                set
                {
                    _minimum = value;
                    Value = _value;
                }
            }

            public decimal Maximum
            {
                get { return _maximum; }
                set
                {
                    _maximum = value;
                    Value = _value;
                }
            }

            public decimal Increment { get; set; }

            public int DecimalPlaces
            {
                get { return _decimalPlaces; }
                set
                {
                    _decimalPlaces = Math.Max(0, value);
                    NormalizeText();
                }
            }

            public decimal Value
            {
                get
                {
                    decimal parsed;
                    return decimal.TryParse(_textBox.Text, NumberStyles.Number, CultureInfo.InvariantCulture, out parsed)
                        ? Clamp(parsed)
                        : _value;
                }
                set
                {
                    _value = Clamp(value);
                    _textBox.Text = _value.ToString("F" + _decimalPlaces, CultureInfo.InvariantCulture);
                }
            }

            private decimal Clamp(decimal value)
            {
                return Math.Max(_minimum, Math.Min(_maximum, value));
            }

            private void NormalizeText()
            {
                _value = Value;
                _textBox.Text = _value.ToString("F" + _decimalPlaces, CultureInfo.InvariantCulture);
            }

            private void TextBoxKeyPress(object sender, KeyPressEventArgs e)
            {
                if (char.IsControl(e.KeyChar) || char.IsDigit(e.KeyChar) || e.KeyChar == '.')
                {
                    return;
                }

                e.Handled = true;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                var bounds = new Rectangle(0, 0, Width - 1, Height - 1);
                using (var path = CreatePath(bounds, 6))
                using (var pen = new Pen(_focused ? AccentColor : LineColor, _focused ? 1.4F : 1F))
                {
                    e.Graphics.DrawPath(pen, path);
                }
            }
        }

        private sealed class InfoButton : Button
        {
            private bool _hovered;

            public InfoButton()
            {
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                BackColor = SurfaceColor;
                ForeColor = MutedColor;
                Cursor = Cursors.Help;
                TabStop = true;
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                base.OnMouseEnter(e);
                _hovered = true;
                Invalidate();
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                _hovered = false;
                Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                var bounds = new Rectangle(1, 1, Width - 3, Height - 3);
                using (var pen = new Pen(_hovered ? AccentColor : LineColor, 1F))
                {
                    e.Graphics.DrawEllipse(pen, bounds);
                }

                TextRenderer.DrawText(e.Graphics, "?", Font, new Rectangle(0, 0, Width, Height), _hovered ? AccentHoverColor : MutedColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            }
        }

        private static GraphicsPath CreatePath(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            var diameter = radius * 2;
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
