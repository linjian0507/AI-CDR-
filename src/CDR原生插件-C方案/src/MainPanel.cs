using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Xml.Linq;
using AIVectorCore;
using CdrApp = Corel.Interop.VGCore.Application;

namespace AIVectorHelper
{
    public class MainPanel : UserControl
    {
        private readonly string _addonDir;
        private readonly string _configPath;
        private readonly string _historyPath;
        private AppConfig _config;
        private CancellationTokenSource _requestCts;
        private readonly Stopwatch _jobStopwatch = new Stopwatch();
        private readonly DispatcherTimer _jobTimer;
        private string _jobTitle = "";
        private string _jobOutcome = "";

        private ComboBox _profileBox, _modelBox, _imageProfileBox, _imageModelBox, _variantBox, _historyBox, _settingsProfileBox;
        private TextBox _svgPrompt, _styleBox, _paletteBox, _widthBox, _heightBox, _layersBox, _svgCode, _editBox;
        private TextBox _imagePrompt, _profileNameBox, _baseUrlBox, _apiKeyBox;
        private ComboBox _settingsModelBox, _settingsImageModelBox;
        private ComboBox _referenceModeBox, _imageResolutionBox, _apiTypeBox, _stylePresetBox, _variantCountBox, _temperatureBox, _imagePresetBox, _imageRouteBox, _vtracerPresetBox, _vtracerModeBox;
        private WrapPanel _vtracerModeRow;
        private TextBox _imagePresetDescription;
        private CheckBox _visionBox, _noBackgroundBox, _autoLayerBox, _reverseLayerBox, _autoImportBox;
        private TextBox _proxyBox;
        private TextBlock _status, _jobStatus, _svgReferenceLabel, _imageSourceLabel, _imageResultLabel;
        private Image _svgReferencePreview, _imageSourcePreview, _imageResultPreview;
        private StackPanel _layerPanel;
        private WebBrowser _svgPreview;
        private TabControl _tabs;
        private Button _svgGenerateButton, _svgImportButton, _svgEditButton, _svgRefineButton, _svgStopButton,
            _imageGenerateButton, _imageImportButton, _imageTraceButton;

        private string _svgReferenceData = "";
        private string _imageSourceData = "";
        private string _currentSvg = "";
        private string _currentImageData = "";
        private readonly List<SvgVariant> _variants = new List<SvgVariant>();
        private readonly List<SvgHistory> _history = new List<SvgHistory>();

        private sealed class SvgVariant
        {
            public string Label;
            public string Svg;
            public IReadOnlyList<SvgLayer> Layers;
            public override string ToString() { return Label; }
        }

        private sealed class SvgHistory
        {
            public string FilePath;
            public string Prompt;
            public int Layers;
            public override string ToString()
            {
                var text = (Prompt ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
                if (text.Length > 30) text = text.Substring(0, 30) + "…";
                return text + " [" + Layers + "层]";
            }
        }

        private sealed class ImagePreset
        {
            public string Name { get; set; }
            public string Prompt { get; set; }
            public string Description { get; set; }
            public bool TraceOnly { get; set; }
            public override string ToString() { return Name; }
        }

        public MainPanel()
        {
            try
            {
                _addonDir = Path.GetDirectoryName(typeof(MainPanel).Assembly.Location) ?? Path.GetTempPath();
                _configPath = Path.Combine(_addonDir, "config.json");
                _historyPath = Path.Combine(_addonDir, "svg-history.tsv");
                _config = LoadConfig();
                LoadHistory();
                _jobTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
                _jobTimer.Tick += (s, e) => RefreshJobStatus();

                var root = new DockPanel { Margin = new Thickness(8) };
                root.Children.Add(BuildStatus());
                _tabs = new TabControl();
                _tabs.Items.Add(new TabItem { Header = "SVG 矢量", Content = BuildSvgTab() });
                _tabs.Items.Add(new TabItem { Header = "AI 生图", Content = BuildImageTab() });
                _tabs.Items.Add(new TabItem { Header = "模型设置", Content = BuildSettingsTab() });
                root.Children.Add(_tabs);
                Content = root;
                RefreshProfiles();
                Log.W("MainPanel 构造完成 ✓ (完整 SVG/生图工作区已加载)");
            }
            catch (Exception ex)
            {
                Log.W("MainPanel 构造失败: " + ex);
                Content = new TextBlock { Text = "AI 矢量助手加载失败: " + ex.Message, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(8) };
            }
        }

        private UIElement BuildStatus()
        {
            var panel = new StackPanel { Margin = new Thickness(0, 6, 0, 0) };
            _jobStatus = new TextBlock { Text = "作业状态：空闲", TextWrapping = TextWrapping.Wrap, FontWeight = FontWeights.SemiBold };
            _status = new TextBlock { Text = "就绪。", TextWrapping = TextWrapping.Wrap, Opacity = 0.82, Margin = new Thickness(0, 2, 0, 0) };
            panel.Children.Add(_jobStatus);
            panel.Children.Add(_status);
            var border = new Border
            {
                BorderBrush = System.Windows.Media.Brushes.SteelBlue,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(5),
                Child = panel
            };
            DockPanel.SetDock(border, Dock.Bottom);
            return border;
        }

        private UIElement BuildSvgTab()
        {
            var panel = NewPanel();
            panel.AllowDrop = true;
            panel.DragOver += (s, e) => OnFileDragOver(e);
            panel.Drop += (s, e) => OnSvgDrop(e);
            panel.Children.Add(Label("模型"));
            _profileBox = new ComboBox { MinWidth = 240, DisplayMemberPath = "Name" };
            _profileBox.SelectionChanged += OnProfileChanged;
            panel.Children.Add(_profileBox);
            _modelBox = new ComboBox { MinWidth = 240, Margin = new Thickness(0, 3, 0, 5) };
            panel.Children.Add(_modelBox);

            panel.Children.Add(Label("提示词（可选：已有参考图时仅补充要求）"));
            _svgPrompt = MultiLine(_config.LastPrompt, 72);
            _svgPrompt.ToolTip = "未选择参考图时必须填写；选择参考图后可留空，插件会按临摹/风格参考模式直接生成。";
            panel.Children.Add(_svgPrompt);

            var presetRow = new WrapPanel { Margin = new Thickness(0, 4, 0, 4) };
            var template = new ComboBox { Width = 160, ItemsSource = SvgTemplates(), SelectedIndex = 0 };
            template.SelectionChanged += (s, e) => { if (template.SelectedIndex > 0) _svgPrompt.Text = (string)template.SelectedItem; };
            presetRow.Children.Add(template);
            _stylePresetBox = new ComboBox { Width = 100, Margin = new Thickness(5, 0, 0, 0), ItemsSource = SvgStyles(), SelectedIndex = 0 };
            _stylePresetBox.SelectionChanged += (s, e) =>
            {
                if (_styleBox != null && _stylePresetBox.SelectedIndex >= 0)
                    _styleBox.Text = SvgStyles()[_stylePresetBox.SelectedIndex];
            };
            presetRow.Children.Add(_stylePresetBox);
            presetRow.Children.Add(new TextBlock { Text = "方案数：", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(7, 0, 1, 0) });
            _variantCountBox = new ComboBox { Width = 120, Margin = new Thickness(5, 0, 0, 0), ItemsSource = new[] { "1 · 一套方案", "2 · 两套方案", "3 · 三套方案" }, SelectedIndex = ChoiceIndex(_config.VariantCount, 1, 3) };
            _variantCountBox.ToolTip = "一次生成几份候选 SVG：1=一份，2=两份，3=三份。方案数越多，请求次数越多。";
            presetRow.Children.Add(new TextBlock { Text = "创意度：", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(7, 0, 1, 0) });
            _temperatureBox = new ComboBox { Width = 120, Margin = new Thickness(5, 0, 0, 0), ItemsSource = new[] { "0.3 · 严谨", "0.6 · 平衡", "0.9 · 发散" }, SelectedIndex = ChoiceIndex(_config.Creativity, 0.6, 0.3) };
            _temperatureBox.ToolTip = "控制生成发散程度：0.3=严谨稳定，0.6=平衡，0.9=更有创意和变化。";
            presetRow.Children.Add(_variantCountBox);
            presetRow.Children.Add(_temperatureBox);
            var referenceMode = string.Equals(_config.ReferenceMode, "style", StringComparison.OrdinalIgnoreCase) ? 2 : 1;
            _referenceModeBox = new ComboBox { Width = 190, Margin = new Thickness(5, 0, 0, 0), ItemsSource = new[] { "无参考（直接按提示生成）", "临摹还原（尽量一致）", "风格参考（借鉴构图/风格）" }, SelectedIndex = referenceMode };
            presetRow.Children.Add(_referenceModeBox);
            var chooseRef = Button("选择参考图", (s, e) => ChooseImage(true));
            chooseRef.Margin = new Thickness(5, 0, 0, 0);
            presetRow.Children.Add(chooseRef);
            presetRow.Children.Add(Button("清除参考图", (s, e) => ClearSvgReference()));
            presetRow.Children.Add(Button("载入 SVG", (s, e) => ChooseSvgFile()));
            panel.Children.Add(presetRow);
            panel.Children.Add(new TextBlock { Text = "说明：方案数决定一次生成的候选数量；创意度决定输出的稳定性与变化幅度。", TextWrapping = TextWrapping.Wrap, Opacity = 0.72, Margin = new Thickness(0, 0, 0, 3) });
            var dropTip = new Border
            {
                BorderBrush = System.Windows.Media.Brushes.SteelBlue,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(6),
                Margin = new Thickness(0, 2, 0, 4),
                Child = new TextBlock { Text = "⇩ 把图片作为参考图拖到这里；把 .svg 拖到这里可直接载入外部 SVG。" }
            };
            panel.Children.Add(dropTip);
            _svgReferenceLabel = new TextBlock { Text = "未选择参考图", Margin = new Thickness(0, 0, 0, 2) };
            panel.Children.Add(_svgReferenceLabel);
            _svgReferencePreview = Preview();
            panel.Children.Add(DropArea(_svgReferencePreview, "拖入参考图到此处", true));

            var sizeRow = new WrapPanel { Margin = new Thickness(0, 5, 0, 3) };
            _widthBox = SmallField(sizeRow, "宽", _config.SvgWidth.ToString());
            _heightBox = SmallField(sizeRow, "高", _config.SvgHeight.ToString());
            _layersBox = SmallField(sizeRow, "层数", _config.LayerCount);
            panel.Children.Add(sizeRow);
            _styleBox = Field(panel, "风格（可选）", "");
            _paletteBox = Field(panel, "配色（可选）", _config.Palette);
            _noBackgroundBox = new CheckBox { Content = "透明背景", IsChecked = _config.NoBackground, Margin = new Thickness(0, 3, 0, 4) };
            panel.Children.Add(_noBackgroundBox);
            _autoLayerBox = new CheckBox { Content = "导入后自动拆分 CDR 图层", IsChecked = _config.AutoLayer, Margin = new Thickness(0, 1, 0, 2) };
            _reverseLayerBox = new CheckBox { Content = "反转拆分图层顺序", IsChecked = _config.ReverseLayers, Margin = new Thickness(0, 1, 0, 2) };
            _autoImportBox = new CheckBox { Content = "生成完成后自动导入 CDR", Margin = new Thickness(0, 1, 0, 4) };
            panel.Children.Add(_autoLayerBox);
            panel.Children.Add(_reverseLayerBox);
            panel.Children.Add(_autoImportBox);

            var actionRow = new WrapPanel();
            _svgGenerateButton = Button("生成 SVG", OnSvgGenerate);
            _svgEditButton = Button("修改当前 SVG", OnSvgEdit);
            _svgRefineButton = Button("对照参考图修正", OnSvgRefine);
            _svgImportButton = Button("导入勾选图层", OnSvgImport);
            _svgStopButton = Button("停止", (s, e) => CancelRequest());
            actionRow.Children.Add(_svgGenerateButton);
            actionRow.Children.Add(_svgEditButton);
            actionRow.Children.Add(_svgRefineButton);
            actionRow.Children.Add(_svgImportButton);
            actionRow.Children.Add(_svgStopButton);
            panel.Children.Add(actionRow);

            panel.Children.Add(Label("修改要求"));
            _editBox = new TextBox { MinWidth = 240, Margin = new Thickness(0, 0, 0, 5) };
            panel.Children.Add(_editBox);
            panel.Children.Add(Label("方案 / 历史"));
            var historyRow = new WrapPanel();
            _variantBox = new ComboBox { Width = 110 };
            _variantBox.SelectionChanged += (s, e) => SelectVariant(_variantBox.SelectedIndex);
            historyRow.Children.Add(_variantBox);
            _historyBox = new ComboBox { Width = 145, Margin = new Thickness(4, 0, 0, 0) };
            historyRow.Children.Add(_historyBox);
            historyRow.Children.Add(Button("载入", (s, e) => LoadSelectedHistory()));
            historyRow.Children.Add(Button("载入并导入", (s, e) => { LoadSelectedHistory(); OnSvgImport(s, e); }));
            historyRow.Children.Add(Button("删除历史", (s, e) => DeleteSelectedHistory()));
            historyRow.Children.Add(Button("复制代码", (s, e) => Clipboard.SetText(_currentSvg ?? "")));
            panel.Children.Add(historyRow);

            panel.Children.Add(Label("SVG 预览"));
            _svgPreview = new WebBrowser { Height = 180, Margin = new Thickness(0, 2, 0, 4) };
            panel.Children.Add(_svgPreview);
            panel.Children.Add(Label("SVG 代码"));
            _svgCode = MultiLine("", 150);
            _svgCode.IsReadOnly = true;
            panel.Children.Add(_svgCode);
            panel.Children.Add(Label("图层（取消勾选可局部导入）"));
            _layerPanel = new StackPanel { Margin = new Thickness(0, 0, 0, 5) };
            panel.Children.Add(_layerPanel);
            var layerButtons = new WrapPanel();
            layerButtons.Children.Add(Button("全选图层", (s, e) => SetAllLayers(true)));
            layerButtons.Children.Add(Button("全不选", (s, e) => SetAllLayers(false)));
            layerButtons.Children.Add(Button("打开输出文件夹", (s, e) => OpenOutputDirectory()));
            panel.Children.Add(layerButtons);
            var scroll = Scroll(panel);
            SetupDropTarget(scroll, true);
            return scroll;
        }

        private UIElement BuildImageTab()
        {
            var panel = NewPanel();
            panel.AllowDrop = true;
            panel.DragOver += (s, e) => OnFileDragOver(e);
            panel.Drop += (s, e) => OnImageDrop(e);
            panel.Children.Add(Label("生图配置（与 SVG 独立）"));
            _imageProfileBox = new ComboBox { MinWidth = 240, DisplayMemberPath = "Name", Margin = new Thickness(0, 0, 0, 3) };
            _imageProfileBox.SelectionChanged += OnImageProfileChanged;
            panel.Children.Add(_imageProfileBox);
            panel.Children.Add(Label("图像模型（可与 SVG 模型不同；可直接输入模型名）"));
            _imageModelBox = new ComboBox { MinWidth = 240, Margin = new Thickness(0, 0, 0, 4) };
            _imageModelBox.IsEditable = true;
            _imageModelBox.ToolTip = "留空时使用当前档案的生图模型或 SVG 模型。";
            panel.Children.Add(_imageModelBox);
            panel.Children.Add(new TextBlock
            {
                Text = "模型名按中转站实际列表填写；常见示例：gpt-image-1、dall-e-3、gemini-2.5-flash-image。图像接口模型优先走 /images，视觉对话模型优先走 /chat/completions。",
                TextWrapping = TextWrapping.Wrap,
                Opacity = 0.72,
                Margin = new Thickness(0, -1, 0, 4)
            });
            panel.Children.Add(Label("图片提示词 / 编辑要求"));
            _imagePrompt = MultiLine("", 70);
            panel.Children.Add(_imagePrompt);
            panel.Children.Add(Label("处理预设"));
            _imagePresetBox = new ComboBox { MinWidth = 240, ItemsSource = ImagePresets(), DisplayMemberPath = "Name", SelectedIndex = 0 };
            _imagePresetBox.SelectionChanged += (s, e) => UpdateImagePresetDescription();
            panel.Children.Add(_imagePresetBox);
            _imagePresetDescription = new TextBox { IsReadOnly = true, TextWrapping = TextWrapping.Wrap, Height = 66, Margin = new Thickness(0, 1, 0, 3) };
            panel.Children.Add(_imagePresetDescription);
            UpdateImagePresetDescription();
            var srcRow = new WrapPanel { Margin = new Thickness(0, 4, 0, 4) };
            srcRow.Children.Add(Button("选择来源图片", (s, e) => ChooseImage(false)));
            srcRow.Children.Add(Button("从 CDR 选中对象获取", (s, e) => GrabSelectionFromCdr()));
            srcRow.Children.Add(Button("清除来源", (s, e) => ClearImageSource()));
            srcRow.Children.Add(new TextBlock { Text = "分辨率", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(5, 0, 2, 4) });
            _imageResolutionBox = new ComboBox { Width = 80, Margin = new Thickness(5, 0, 0, 0), ItemsSource = new[] { "1k", "2k", "4k" }, SelectedIndex = 1 };
            srcRow.Children.Add(_imageResolutionBox);
            srcRow.Children.Add(new TextBlock { Text = "路由", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(5, 0, 2, 4) });
            _imageRouteBox = new ComboBox { Width = 95, Margin = new Thickness(0, 0, 0, 0), ItemsSource = new[] { "自动路由", "图像接口", "对话接口" }, SelectedIndex = 0, ToolTip = "自动按模型和是否有来源图片选择 /images 或 /chat/completions。" };
            srcRow.Children.Add(_imageRouteBox);
            _imageGenerateButton = Button("生成 / 编辑图片", OnImageGenerate);
            srcRow.Children.Add(_imageGenerateButton);
            panel.Children.Add(srcRow);
            panel.Children.Add(new TextBlock { Text = "路由说明：图像接口适用于 gpt-image、DALL·E、Imagen 等；对话接口适用于会在对话中返回图片链接/数据的视觉模型；自动只按模型选择正确接口，不会把图像模型误发到 Chat Completions。", TextWrapping = TextWrapping.Wrap, Opacity = 0.72, Margin = new Thickness(0, 0, 0, 3) });
            var vtracerRow = new WrapPanel { Margin = new Thickness(0, 0, 0, 3) };
            vtracerRow.Children.Add(new TextBlock { Text = "描摹方案", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0) });
            _vtracerPresetBox = new ComboBox
            {
                Width = 285,
                ItemsSource = new[]
                {
                    "CorelDRAW PowerTRACE（CDR 内置）",
                    "VTracer（离线描摹）"
                },
                SelectedIndex = 0
            };
            _vtracerPresetBox.SelectionChanged += (s, e) => UpdateTraceSchemeUi();
            vtracerRow.Children.Add(_vtracerPresetBox);
            panel.Children.Add(vtracerRow);
            _vtracerModeRow = new WrapPanel { Margin = new Thickness(0, 0, 0, 3) };
            _vtracerModeRow.Children.Add(new TextBlock { Text = "VTracer 预设", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 4, 0) });
            _vtracerModeBox = new ComboBox
            {
                Width = 285,
                ItemsSource = new[]
                {
                    "默认：彩色 Logo / 扁平插画",
                    "黑白线稿 / 扫描件",
                    "照片 / 多色图 / 油画素材"
                },
                SelectedIndex = 0
            };
            _vtracerModeRow.Children.Add(_vtracerModeBox);
            panel.Children.Add(_vtracerModeRow);
            panel.Children.Add(new TextBlock
            {
                Text = "两种描摹引擎可自由切换：PowerTRACE 直接调用 CDR 内置描摹；VTracer 选中后可继续选择默认、黑白或照片/油画类预设。",
                TextWrapping = TextWrapping.Wrap,
                Opacity = 0.72,
                Margin = new Thickness(0, 0, 0, 3)
            });
            _imageSourceLabel = new TextBlock { Text = "未选择来源图片：直接生图；选择后执行图片编辑。", TextWrapping = TextWrapping.Wrap };
            panel.Children.Add(_imageSourceLabel);
            _imageSourcePreview = Preview();
            panel.Children.Add(DropArea(_imageSourcePreview, "拖入来源图片到此处", false));
            panel.Children.Add(Label("生成结果"));
            _imageResultPreview = Preview();
            panel.Children.Add(_imageResultPreview);
            _imageResultLabel = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 2, 0, 4) };
            panel.Children.Add(_imageResultLabel);
            var resultRow = new WrapPanel();
            _imageImportButton = Button("导入图片到 CDR", OnImageImport);
            _imageTraceButton = Button("PowerTRACE 描摹并导入", OnImageTrace);
            resultRow.Children.Add(_imageImportButton);
            resultRow.Children.Add(_imageTraceButton);
            resultRow.Children.Add(Button("保存当前图片", (s, e) => SaveCurrentImage()));
            resultRow.Children.Add(Button("作为 SVG 参考图", (s, e) => UseImageAsSvgReference()));
            panel.Children.Add(resultRow);
            UpdateTraceSchemeUi();
            var scroll = Scroll(panel);
            SetupDropTarget(scroll, false);
            return scroll;
        }

        private UIElement BuildSettingsTab()
        {
            var panel = NewPanel();
            panel.Children.Add(Label("模型档案"));
            var profileRow = new WrapPanel();
            _settingsProfileBox = new ComboBox { Width = 165, DisplayMemberPath = "Name" };
            _settingsProfileBox.SelectionChanged += (s, e) =>
            {
                var index = _settingsProfileBox.SelectedIndex;
                if (index >= 0 && index < _config.Profiles.Count && _profileNameBox != null)
                {
                    _config.ActiveIndex = index;
                    if (_profileBox != null) _profileBox.SelectedIndex = index;
                    FillSettings(_config.Profiles[index]);
                }
            };
            profileRow.Children.Add(_settingsProfileBox);
            profileRow.Children.Add(Button("新增", (s, e) =>
            {
                _config.Profiles.Add(new ApiProfile { Name = "新模型", ApiType = "openai" });
                var index = _config.Profiles.Count - 1;
                _config.ActiveIndex = index;
                RefreshProfiles(index);
                FillSettings(_config.Profiles[index]);
                SaveConfig();
            }));
            profileRow.Children.Add(Button("删除", (s, e) =>
            {
                var index = _settingsProfileBox.SelectedIndex;
                if (index >= 0 && _config.Profiles.Count > 1)
                {
                    _config.Profiles.RemoveAt(index);
                    _config.ActiveIndex = Math.Max(0, Math.Min(index, _config.Profiles.Count - 1));
                    RefreshProfiles(_config.ActiveIndex);
                    FillSettings(_config.Profiles[_config.ActiveIndex]);
                    SaveConfig();
                }
            }));
            profileRow.Children.Add(Button("复制", (s, e) =>
            {
                var index = _settingsProfileBox.SelectedIndex;
                if (index < 0 || index >= _config.Profiles.Count) return;
                var copy = _config.Profiles[index].Clone();
                copy.Name = (copy.Name ?? "模型") + " 副本";
                _config.Profiles.Add(copy);
                var copyIndex = _config.Profiles.Count - 1;
                _config.ActiveIndex = copyIndex;
                RefreshProfiles(copyIndex);
                FillSettings(copy);
                SaveConfig();
            }));
            panel.Children.Add(profileRow);

            _profileNameBox = Field(panel, "名称", "");
            _apiTypeBox = new ComboBox { ItemsSource = new[] { "openai", "anthropic" }, Margin = new Thickness(0, 0, 0, 4) };
            panel.Children.Add(Label("接口类型")); panel.Children.Add(_apiTypeBox);
            _baseUrlBox = Field(panel, "API 地址", "");
            panel.Children.Add(Label("SVG 模型"));
            _settingsModelBox = new ComboBox { MinWidth = 240, IsEditable = true, Margin = new Thickness(0, 0, 0, 3), ToolTip = "点击选择已获取的模型；也可直接输入模型名称。" };
            panel.Children.Add(_settingsModelBox);
            panel.Children.Add(Label("生图模型（留空则使用 SVG 模型）"));
            _settingsImageModelBox = new ComboBox { MinWidth = 240, IsEditable = true, Margin = new Thickness(0, 0, 0, 3), ToolTip = "点击选择已获取的模型；也可直接输入模型名称。" };
            panel.Children.Add(_settingsImageModelBox);
            panel.Children.Add(Label("API Key"));
            _apiKeyBox = new TextBox { MinWidth = 240, Margin = new Thickness(0, 0, 0, 4) };
            panel.Children.Add(_apiKeyBox);
            _visionBox = new CheckBox { Content = "支持参考图", Margin = new Thickness(0, 2, 0, 6) };
            panel.Children.Add(_visionBox);
            panel.Children.Add(Button("保存模型设置", (s, e) => SaveSettings(_settingsProfileBox.SelectedIndex)));
            panel.Children.Add(Button("获取模型列表", async (s, e) =>
            {
                SyncSettingsToProfile(_settingsProfileBox.SelectedIndex);
                await FetchModelsAsync(_settingsProfileBox.SelectedIndex);
            }));
            panel.Children.Add(Button("测试中转连接", async (s, e) =>
            {
                SyncSettingsToProfile(_settingsProfileBox.SelectedIndex);
                await TestApiConnectionAsync(_settingsProfileBox.SelectedIndex);
            }));
            _proxyBox = Field(panel, "代理（可选）", _config.Proxy);
            panel.Children.Add(Button("保存网络设置", (s, e) =>
            {
                _config.Proxy = _proxyBox.Text.Trim();
                SaveConfig();
                SetStatus("网络设置已保存。");
            }));
            RefreshProfiles();
            if (_config.ActiveProfile != null) FillSettings(_config.ActiveProfile);
            else RefreshSettingsModelSelectors(null);
            return Scroll(panel);
        }

        private static StackPanel NewPanel() { return new StackPanel { Margin = new Thickness(8) }; }
        private static ScrollViewer Scroll(UIElement child) { return new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = child }; }
        private static TextBlock Label(string text) { return new TextBlock { Text = text, Margin = new Thickness(0, 4, 0, 2) }; }
        private static TextBox MultiLine(string value, double height) { return new TextBox { Text = value ?? "", MinWidth = 240, Height = height, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(0, 0, 0, 3) }; }
        private static Image Preview() { return new Image { MaxWidth = 240, MaxHeight = 150, Stretch = System.Windows.Media.Stretch.Uniform, Margin = new Thickness(0, 2, 0, 4) }; }
        private Border DropArea(Image preview, string hint, bool svgReference)
        {
            var border = new Border
            {
                BorderThickness = new Thickness(1),
                BorderBrush = System.Windows.Media.Brushes.Gray,
                Padding = new Thickness(5),
                AllowDrop = true,
                Margin = new Thickness(0, 2, 0, 4)
            };
            var stack = new StackPanel();
            stack.Children.Add(new TextBlock { Text = hint, Opacity = 0.7, TextWrapping = TextWrapping.Wrap });
            stack.Children.Add(preview);
            border.Child = stack;
            border.DragEnter += (s, e) =>
            {
                e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
                e.Handled = true;
            };
            border.Drop += (s, e) =>
            {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files == null || files.Length == 0) return;
                HandleDroppedFile(files[0], svgReference);
                e.Handled = true;
            };
            return border;
        }
        private static Button Button(string text, RoutedEventHandler handler) { return new Button { Content = text, Padding = new Thickness(7, 3, 7, 3), Margin = new Thickness(0, 0, 4, 4) }.With(handler); }
        private static TextBox Field(Panel panel, string label, string value) { panel.Children.Add(Label(label)); var box = new TextBox { Text = value ?? "", MinWidth = 240, Margin = new Thickness(0, 0, 0, 3) }; panel.Children.Add(box); return box; }
        private static TextBox SmallField(Panel panel, string label, string value) { panel.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0, 0, 2, 0), VerticalAlignment = VerticalAlignment.Center }); var box = new TextBox { Text = value ?? "", Width = 48, Margin = new Thickness(0, 0, 5, 0) }; panel.Children.Add(box); return box; }

        private void RefreshProfiles(int? preferredIndex = null)
        {
            if (_config.Profiles == null) _config.Profiles = new List<ApiProfile>();
            var index = preferredIndex.HasValue
                ? preferredIndex.Value
                : _config.ActiveIndex;
            index = _config.Profiles.Count == 0
                ? -1
                : Math.Max(0, Math.Min(index, _config.Profiles.Count - 1));

            if (_profileBox != null)
            {
                _profileBox.ItemsSource = null;
                _profileBox.ItemsSource = _config.Profiles;
                _profileBox.SelectedIndex = index;
            }
            if (_imageProfileBox != null)
            {
                var imageIndex = preferredIndex.HasValue ? index : _imageProfileBox.SelectedIndex;
                if (imageIndex < 0 || imageIndex >= _config.Profiles.Count)
                {
                    imageIndex = _config.Profiles.FindIndex(x =>
                        (x.Models ?? new List<string>()).Any(m => m.IndexOf("image", StringComparison.OrdinalIgnoreCase) >= 0) ||
                        (x.Model ?? "").IndexOf("image", StringComparison.OrdinalIgnoreCase) >= 0);
                }
                if (imageIndex < 0) imageIndex = index;
                _imageProfileBox.ItemsSource = null;
                _imageProfileBox.ItemsSource = _config.Profiles;
                _imageProfileBox.SelectedIndex = imageIndex;
            }
            if (_settingsProfileBox != null)
            {
                _settingsProfileBox.ItemsSource = null;
                _settingsProfileBox.ItemsSource = _config.Profiles;
                _settingsProfileBox.SelectedIndex = index;
            }
        }

        private void OnProfileChanged(object sender, SelectionChangedEventArgs e)
        {
            var p = ActiveProfile;
            if (p == null) return;
            if (_profileBox != null && _profileBox.SelectedIndex >= 0)
                _config.ActiveIndex = _profileBox.SelectedIndex;
            var models = new List<string>(p.Models ?? new List<string>());
            if (!string.IsNullOrWhiteSpace(p.Model) && !models.Contains(p.Model)) models.Insert(0, p.Model);
            _modelBox.ItemsSource = models;
            _modelBox.SelectedItem = p.Model;
        }

        private void OnImageProfileChanged(object sender, SelectionChangedEventArgs e)
        {
            var p = ActiveImageProfile;
            if (p == null || _imageModelBox == null) return;
            var models = new List<string>(p.Models ?? new List<string>());
            var preferred = string.IsNullOrWhiteSpace(p.ImageModel) ? p.Model : p.ImageModel;
            if (!string.IsNullOrWhiteSpace(preferred) && !models.Contains(preferred)) models.Insert(0, preferred);
            _imageModelBox.ItemsSource = models;
            _imageModelBox.SelectedItem = preferred;
        }

        private ApiProfile ActiveProfile { get { return _profileBox == null ? _config.ActiveProfile : _profileBox.SelectedItem as ApiProfile; } }
        private ApiProfile ActiveImageProfile { get { return _imageProfileBox == null ? ActiveProfile : _imageProfileBox.SelectedItem as ApiProfile; } }

        private async void OnSvgGenerate(object sender, RoutedEventArgs e)
        {
            await RunSvg(async generator =>
            {
                int count = 1;
                int.TryParse(ChoiceValue(_variantCountBox == null ? "" : _variantCountBox.SelectedItem as string, "1"), out count);
                count = Math.Max(1, Math.Min(3, count));
                for (var i = 0; i < count; i++)
                {
                    var result = await generator.GenerateAsync(ActiveProfile, SvgOptions(), _svgPrompt.Text, SvgReference(), Token, Delta);
                    ApplySvg(result, "方案 " + (_variants.Count + 1));
                }
            });
        }

        private async void OnSvgEdit(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_currentSvg)) { SetStatus("请先生成 SVG。"); return; }
            await RunSvg(async generator =>
            {
                var result = await generator.EditAsync(ActiveProfile, SvgOptions(), _currentSvg, _editBox.Text, Token, Delta);
                ApplySvg(result, "修改 " + (_variants.Count + 1));
            });
        }

        private async void OnSvgRefine(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_currentSvg)) { SetStatus("请先生成 SVG。"); return; }
            await RunSvg(async generator =>
            {
                var result = await generator.RefineAgainstReferenceAsync(ActiveProfile, SvgOptions(), _currentSvg, SvgReference(), Token, Delta);
                ApplySvg(result, "修正 " + (_variants.Count + 1));
            });
        }

        private async Task RunSvg(Func<SvgGenerator, Task> action)
        {
            if (ActiveProfile == null) { SetStatus("请先在模型设置中添加模型档案。"); return; }
            if (string.IsNullOrWhiteSpace(_svgPrompt.Text) && string.IsNullOrWhiteSpace(_currentSvg) && string.IsNullOrWhiteSpace(_svgReferenceData))
            {
                SetStatus("请输入 SVG 提示词或选择参考图。");
                return;
            }
            BeginBusy("SVG 请求中...");
            try
            {
                using (var client = new ApiClient(_proxyBox == null ? _config.Proxy : _proxyBox.Text.Trim()))
                {
                    await action(new SvgGenerator(client));
                    SaveConfig();
                }
            }
            catch (OperationCanceledException) { MarkJobOutcome("已停止"); SetStatus("SVG 任务已停止。"); }
            catch (Exception ex) { MarkJobOutcome("失败"); SetStatus("SVG 失败: " + ex.Message); Log.W("SVG 失败: " + ex); }
            finally { EndBusy(); }
        }

        private void ApplySvg(SvgGenerationResult result, string label, bool allowAutoImport = true)
        {
            _currentSvg = result.Svg;
            _svgCode.Text = result.Svg;
            UpdateSvgPreview(result.Svg);
            _variants.Add(new SvgVariant { Label = label, Svg = result.Svg, Layers = result.Layers });
            _variantBox.ItemsSource = null; _variantBox.ItemsSource = _variants; _variantBox.SelectedIndex = _variants.Count - 1;
            RenderLayers(result.Layers);
            SaveSvgHistory(result);
            SetStatus("SVG 成功 ✓ " + result.Layers.Count + " 个图层，可勾选后导入 CDR。");
            Log.W("SVG 成功: length=" + result.Svg.Length + " layers=" + result.Layers.Count);
            if (allowAutoImport && _autoImportBox != null && _autoImportBox.IsChecked == true)
                OnSvgImport(this, new RoutedEventArgs());
        }

        private void SelectVariant(int index)
        {
            if (index < 0 || index >= _variants.Count) return;
            var variant = _variants[index];
            _currentSvg = variant.Svg; _svgCode.Text = variant.Svg; UpdateSvgPreview(variant.Svg); RenderLayers(variant.Layers);
        }

        private void RenderLayers(IReadOnlyList<SvgLayer> layers)
        {
            _layerPanel.Children.Clear();
            if (layers == null || layers.Count == 0) { _layerPanel.Children.Add(new TextBlock { Text = "未检测到分组图层，导入完整 SVG。" }); return; }
            foreach (var layer in layers)
                _layerPanel.Children.Add(new CheckBox { Content = string.IsNullOrWhiteSpace(layer.Name) ? "图层 " + (layer.Index + 1) : layer.Name, IsChecked = true, Tag = layer });
        }

        private void OnSvgImport(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_currentSvg)) { SetStatus("没有可导入的 SVG。"); return; }
            try
            {
                var svg = FilterSelectedLayers(_currentSvg);
                var path = WriteTempFile("svg", ".svg", Encoding.UTF8.GetBytes(svg));
                var app = GetApp(out var via);
                if (app == null) throw new InvalidOperationException("没有取得 CorelDRAW Application。");
                var doc = app.ActiveDocument ?? app.CreateDocument();
                var importedSvg = ImportFileToActiveLayer(doc, path);
                UnlockImportedContent(doc, importedSvg);
                var made = 0;
                if (_autoLayerBox != null && _autoLayerBox.IsChecked == true)
                {
                    try
                    {
                        dynamic imported = doc.ActiveShape;
                        made = SplitImportedLayers(app, doc, imported, _layerPanel.Children.OfType<CheckBox>().Select(x => ((SvgLayer)x.Tag).Name).ToList());
                    }
                    catch (Exception splitError) { Log.W("自动拆分图层失败: " + splitError); }
                }
                UnlockImportedContent(doc, doc.ActiveShape);
                SetStatus("SVG 导入成功 ✓ " + via + (made > 0 ? "，已拆分 " + made + " 个 CDR 图层。" : ""));
                Log.W("SVG 导入成功: " + path);
            }
            catch (Exception ex) { SetStatus("SVG 导入失败: " + ex.Message); Log.W("SVG 导入失败: " + ex); }
        }

        private int SplitImportedLayers(object appObject, object docObject, dynamic top, IReadOnlyList<string> names)
        {
            if (top == null) return 0;
            try
            {
                if ((int)top.Type != 7) return 0;
                dynamic container = top;
                dynamic children = container.Shapes;
                var count = (int)children.Count;
                if (count < 2) return 0;
                var shapes = new List<dynamic>();
                for (var i = 1; i <= count; i++) shapes.Add(children.Item(i));
                container.Ungroup();
                if (_reverseLayerBox != null && _reverseLayerBox.IsChecked == true) shapes.Reverse();
                dynamic doc = docObject;
                dynamic page = doc.ActivePage;
                var made = 0;
                for (var i = shapes.Count - 1; i >= 0; i--)
                {
                    var nameIndex = shapes.Count - 1 - i;
                    var name = names != null && nameIndex < names.Count && !string.IsNullOrWhiteSpace(names[nameIndex])
                        ? names[nameIndex] : "AI图层 " + (made + 1);
                    dynamic layer = page.CreateLayer(name);
                    UnlockCdrObject(layer);
                    shapes[i].MoveToLayer(layer);
                    UnlockImportedContent(doc, shapes[i]);
                    made++;
                }
                return made;
            }
            catch { return 0; }
        }

        private string FilterSelectedLayers(string svg)
        {
            var enabled = _layerPanel.Children.OfType<CheckBox>().Where(x => x.IsChecked == true).Select(x => (SvgLayer)x.Tag).ToList();
            if (enabled.Count == 0 && _layerPanel.Children.OfType<CheckBox>().Any()) throw new InvalidOperationException("至少勾选一个图层。");
            if (!enabled.Any()) return svg;
            var ids = new HashSet<string>(enabled.Select(x => x.Id));
            var names = new HashSet<string>(enabled.Select(x => x.Name));
            var document = XDocument.Parse(svg, LoadOptions.PreserveWhitespace);
            foreach (var group in document.Root.Elements().Where(x => x.Name.LocalName == "g").ToList())
            {
                var id = (string)group.Attribute("id") ?? "";
                var name = (string)group.Attribute("data-name") ?? "";
                if (!ids.Contains(id) && !names.Contains(name)) group.Remove();
            }
            return document.ToString(SaveOptions.DisableFormatting);
        }

        private async void OnImageGenerate(object sender, RoutedEventArgs e)
        {
            var preset = _imagePresetBox == null ? null : _imagePresetBox.SelectedItem as ImagePreset;
            if (preset != null && preset.TraceOnly)
            {
                if (string.IsNullOrWhiteSpace(_imageSourceData)) { SetStatus("描摹预设需要先载入来源图片。"); return; }
                OnImageTrace(sender, e);
                return;
            }
            var p = ActiveImageProfile;
            if (p == null) { SetStatus("请先选择模型档案。"); return; }
            if (string.IsNullOrWhiteSpace(_imagePrompt.Text) && (preset == null || string.IsNullOrWhiteSpace(preset.Prompt)) && string.IsNullOrWhiteSpace(_imageSourceData))
            {
                SetStatus("请输入图片提示词、选择处理预设或载入来源图片。");
                return;
            }
            BeginBusy("生图请求中...");
            try
            {
                using (var client = new ApiClient(_proxyBox == null ? _config.Proxy : _proxyBox.Text.Trim()))
                {
                    var image = new ImageClient(client);
                    var model = !string.IsNullOrWhiteSpace(_imageModelBox.Text) ? _imageModelBox.Text.Trim() : (p.ImageModel ?? p.Model);
                    var prompt = ((preset == null ? "" : preset.Prompt) + "\n" + (_imagePrompt.Text ?? "")).Trim();
                    if (string.IsNullOrWhiteSpace(prompt)) prompt = "请处理这张图片。";
                    var selectedRoute = _imageRouteBox == null ? "自动路由" : (_imageRouteBox.SelectedItem as string ?? "自动路由");
                    var hasSource = !string.IsNullOrWhiteSpace(_imageSourceData);
                    var looksImageModel = LooksLikeImageModel(model);
                    if (selectedRoute == "对话接口" && looksImageModel)
                        throw new InvalidOperationException("当前模型是图像生成模型，不能走 Chat Completions。请改为“自动路由”或“图像接口”。");
                    if (selectedRoute == "图像接口" && !looksImageModel)
                        throw new InvalidOperationException("当前模型看起来是对话模型，不能走图像接口。请改为“自动路由”或“对话接口”。");
                    var useChat = selectedRoute == "对话接口" || (selectedRoute == "自动路由" && !looksImageModel);
                    var routeLabel = useChat ? "Chat Completions" : "Images API";
                    SetStatus("生图请求中 · 模型 " + model + " · " + routeLabel);
                    _imageResultLabel.Text = "正在发送实际提示词：" + prompt;
                    Log.W("生图请求: model=" + model + " route=" + routeLabel + " source=" + hasSource
                        + " promptLength=" + prompt.Length + " prompt=" + prompt.Replace("\r", " ").Replace("\n", " "));
                    var result = useChat
                        ? await image.ChatAsync(p, model, prompt, _imageSourceData, (string)_imageResolutionBox.SelectedItem, Token)
                        : (hasSource
                            ? await image.EditAsync(p, model, prompt, _imageSourceData, (string)_imageResolutionBox.SelectedItem, Token)
                            : await image.GenerateAsync(p, model, prompt, (string)_imageResolutionBox.SelectedItem, Token));
                    _currentImageData = await ToDataUrl(result);
                    SetPreview(_imageResultPreview, _currentImageData);
                    var bytes = DecodeDataUrl(_currentImageData);
                    var file = WriteTempFile("image", ".png", bytes);
                    _imageResultLabel.Text = "生图成功 ✓ 已保存: " + file + "\n实际提示词：" + prompt;
                    SetStatus("生图成功 ✓ 可导入 CDR 或用 VTracer 本地描摹。");
                    Log.W("生图成功: bytes=" + bytes.Length);
                }
            }
            catch (OperationCanceledException) { MarkJobOutcome("已停止"); SetStatus("生图任务已停止。"); }
            catch (Exception ex) { MarkJobOutcome("失败"); SetStatus("生图失败: " + ex.Message); Log.W("生图失败: " + ex); }
            finally { EndBusy(); }
        }

        private static bool LooksLikeImageModel(string model)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(
                model ?? "",
                "gpt[-_]?image|dall[-_]?e|image[-_]?\\d|image[-_]preview|imagen|gemini.*image|flux|seedream",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        private void OnImageImport(object sender, RoutedEventArgs e)
        {
            try
            {
                var path = CurrentImageFile();
                var app = GetApp(out var via);
                if (app == null) throw new InvalidOperationException("没有取得 CorelDRAW Application。");
                var doc = app.ActiveDocument ?? app.CreateDocument();
                var importedImage = ImportFileToActiveLayer(doc, path);
                UnlockImportedContent(doc, importedImage);
                SetStatus("图片导入成功 ✓ " + via);
            }
            catch (Exception ex) { SetStatus("图片导入失败: " + ex.Message); Log.W("图片导入失败: " + ex); }
        }

        private static void UnlockImportedContent(dynamic doc, dynamic root)
        {
            try { UnlockCdrObject(doc == null ? null : doc.ActiveLayer); } catch { }
            if (root == null) return;
            UnlockCdrObject(root);
            try
            {
                dynamic children = root.Shapes;
                var count = (int)children.Count;
                for (var i = 1; i <= count; i++)
                    UnlockImportedContent(doc, children.Item(i));
            }
            catch { }
        }

        private static dynamic ImportFileToActiveLayer(dynamic doc, string path)
        {
            if (doc == null) throw new InvalidOperationException("没有取得 CorelDRAW 文档。");
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException("待导入文件不存在。", path);

            // 使用 CDR 当前运行时的动态 Import(path) 默认签名，
            // 避免把旧版 Corel.Interop.VGCore 的枚举和 StructImportOptions
            // 强行传给 CorelDRAW 27 等新版本，降低跨版本崩溃风险。
            dynamic layer = doc.ActiveLayer;
            layer.Import(path);
            return doc.ActiveShape;
        }

        private static void UnlockCdrObject(dynamic value)
        {
            if (value == null) return;
            try { value.Locked = false; } catch { }
            try { value.Editable = true; } catch { }
            try { value.AllowEdit = true; } catch { }
        }

        private async void OnImageTrace(object sender, RoutedEventArgs e)
        {
            var schemeIndex = TraceSchemeIndex();
            var schemeName = TraceSchemeName();
            BeginBusy(schemeIndex == 0
                ? "CorelDRAW PowerTRACE 描摹中..."
                : "VTracer 本地描摹中...");
            try
            {
                var input = CurrentTraceImageFile();
                if (schemeIndex == 0)
                {
                    // PowerTRACE 是 CDR 的进程内 COM 对象模型，只能在当前 UI/COM 线程调用。
                    RunPowerTrace(input);
                    _imageResultLabel.Text = "CorelDRAW PowerTRACE 完成 ✓ 已导入当前文档。来源文件: " + input;
                    SetStatus("CorelDRAW PowerTRACE 描摹完成 ✓ 矢量已导入当前文档。");
                    Log.W("PowerTRACE 描摹完成: input=" + input);
                    return;
                }

                var output = WriteTempFile("vtracer", ".svg", new byte[0]);
                File.Delete(output);
                string commandOutput;
                commandOutput = await RunVTracerAsync(input, output, VTracerArguments());
                if (!File.Exists(output)) throw new InvalidOperationException("VTracer 未生成 SVG 文件。");
                var svg = File.ReadAllText(output, Encoding.UTF8);
                if (string.IsNullOrWhiteSpace(SvgText.Extract(svg))) throw new InvalidOperationException(
                    "VTracer 输出不是有效 SVG。" + Environment.NewLine + commandOutput);
                var result = new SvgGenerationResult { Svg = svg, Layers = SvgLayers.Parse(svg) };
                ApplySvg(result, schemeName, false);
                _imageResultLabel.Text = schemeName + " 完成 ✓ 已输出: " + output;
                if (_tabs != null) _tabs.SelectedIndex = 0;
                OnSvgImport(sender, e);
                Log.W(schemeName + " 描摹完成: input=" + input + " output=" + output + " " + commandOutput);
            }
            catch (OperationCanceledException) { MarkJobOutcome("已停止"); SetStatus("描摹任务已停止。"); }
            catch (Exception ex)
            {
                MarkJobOutcome("失败");
                SetStatus(schemeName + " 描摹失败: " + ex.Message);
                Log.W(schemeName + " 描摹失败: " + ex);
            }
            finally { EndBusy(); }
        }

        private bool IsPowerTraceSelected()
        {
            return TraceSchemeIndex() == 0;
        }

        private int TraceSchemeIndex()
        {
            return _vtracerPresetBox == null ? 0 : Math.Max(0, _vtracerPresetBox.SelectedIndex);
        }

        private string TraceSchemeName()
        {
            switch (TraceSchemeIndex())
            {
                case 1: return "VTracer";
                default: return "CorelDRAW PowerTRACE";
            }
        }

        private void UpdateTraceSchemeUi()
        {
            if (_imageTraceButton != null)
                _imageTraceButton.Content = TraceSchemeName() + " 描摹并导入";
            if (_vtracerModeRow != null)
                _vtracerModeRow.Visibility = TraceSchemeIndex() == 1
                    ? Visibility.Visible
                    : Visibility.Collapsed;
        }

        // 使用 CDR 原生对象模型调用内置 PowerTRACE：
        // Bitmap.Trace(3) = cdrTraceClipart，失败时回退到 Trace(5) = cdrTraceHighQualityImage，
        // 再由 TraceSettings.Finish() 生成可编辑矢量。
        private static dynamic TraceBitmapShape(dynamic shape)
        {
            if (shape == null) return null;
            dynamic bitmap = null;
            try { bitmap = shape.Bitmap; } catch { }
            if (bitmap == null) return null;

            dynamic settings = null;
            try { settings = bitmap.Trace(3); } catch { }
            if (settings == null)
            {
                try { settings = bitmap.Trace(5); } catch { }
            }
            if (settings == null) return null;

            try { settings.DeleteOriginalObject = true; } catch { }
            dynamic traced = null;
            try { traced = settings.Finish(); } catch { }
            return traced;
        }

        private void RunPowerTrace(string input)
        {
            var app = GetApp(out var via);
            if (app == null) throw new InvalidOperationException("没有取得 CorelDRAW Application。");

            try { app.Visible = true; } catch { }

            dynamic workDoc = null;
            try { workDoc = app.ActiveDocument; } catch { }
            if (workDoc == null)
            {
                workDoc = app.CreateDocument();
                Log.W("PowerTRACE: CorelDRAW 中没有活动文档，已新建文档。");
            }

            dynamic tempDoc = null;
            try
            {
                tempDoc = app.OpenDocument(input);
                dynamic page = tempDoc.ActivePage;
                dynamic shapes = page.Shapes;
                var count = (int)shapes.Count;
                if (count <= 0) throw new InvalidOperationException("临时文档中没有可描摹的位图对象。");

                for (var i = 1; i <= count; i++)
                {
                    try { shapes.Item(i).Locked = false; } catch { }
                }

                dynamic source = count == 1 ? shapes.Item(1) : shapes.All().Group();
                UnlockCdrObject(source);
                var traced = TraceBitmapShape(source);
                if (traced == null) throw new InvalidOperationException("PowerTRACE 未返回描摹结果。");
                UnlockCdrObject(traced);

                // 必须先复制，再切回工作文档粘贴，最后关闭临时文档。
                // 这是跨文档自动化下更稳定的复制、粘贴和关闭顺序。
                traced.Copy();
                try { workDoc.Activate(); } catch { }
                dynamic pasted = workDoc.ActiveLayer.Paste();
                UnlockImportedContent(workDoc, pasted);
                try { app.Refresh(); } catch { }
                Log.W("PowerTRACE: 临时文档描摹并粘贴回当前文档 ✓ via=" + via);
            }
            finally
            {
                if (tempDoc != null)
                {
                    try { tempDoc.Dirty = false; } catch { }
                    try { tempDoc.Close(); } catch (Exception ex) { Log.W("PowerTRACE: 关闭临时文档失败: " + ex.Message); }
                }
                try { workDoc.Activate(); } catch { }
                try { app.ActiveWindow.Activate(); } catch { }
            }
        }

        private string VTracerExecutablePath()
        {
            return Path.Combine(_addonDir, "tools", "vtracer", "vtracer.exe");
        }

        private string CurrentTraceImageFile()
        {
            var data = !string.IsNullOrWhiteSpace(_currentImageData) ? _currentImageData : _imageSourceData;
            if (string.IsNullOrWhiteSpace(data)) throw new InvalidOperationException("请先载入来源图片或生成图片。");
            var extension = data.StartsWith("data:image/jpeg", StringComparison.OrdinalIgnoreCase) ? ".jpg"
                : data.StartsWith("data:image/webp", StringComparison.OrdinalIgnoreCase) ? ".webp"
                : data.StartsWith("data:image/bmp", StringComparison.OrdinalIgnoreCase) ? ".bmp" : ".png";
            return WriteTempFile("vtracer-input", extension, DecodeDataUrl(data));
        }

        private string VTracerPresetName()
        {
            return _vtracerModeBox == null
                ? "VTracer：默认彩色 Logo / 扁平插画"
                : (_vtracerModeBox.SelectedItem as string ?? "VTracer：默认彩色 Logo / 扁平插画");
        }

        private string VTracerArguments()
        {
            switch (_vtracerModeBox == null ? 0 : _vtracerModeBox.SelectedIndex)
            {
                case 1: return "--preset bw --filter-speckle 2 --optimize 2 --path-precision 2";
                case 2: return "--preset photo --max-colors 24 --filter-speckle 8 --optimize 2 --path-precision 2";
                default: return "--preset poster --max-colors 12 --filter-speckle 4 --optimize 2 --path-precision 2";
            }
        }

        private async Task<string> RunVTracerAsync(string input, string output, string options)
        {
            var executable = VTracerExecutablePath();
            if (!File.Exists(executable)) throw new FileNotFoundException("未找到 VTracer 程序，请检查插件 tools\\vtracer 目录。", executable);
            var arguments = "--input " + QuoteArgument(input) + " --output " + QuoteArgument(output) + " " + options;
            return await Task.Run(async () =>
            {
                using (var process = new Process())
                {
                    process.StartInfo = new ProcessStartInfo
                    {
                        FileName = executable,
                        Arguments = arguments,
                        WorkingDirectory = Path.GetDirectoryName(executable),
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };
                    if (!process.Start()) throw new InvalidOperationException("VTracer 进程未能启动。");
                    var standardOutput = process.StandardOutput.ReadToEndAsync();
                    var standardError = process.StandardError.ReadToEndAsync();
                    await Task.WhenAll(standardOutput, standardError);
                    process.WaitForExit();
                    var log = (standardOutput.Result + Environment.NewLine + standardError.Result).Trim();
                    if (process.ExitCode != 0) throw new InvalidOperationException("VTracer 退出码 " + process.ExitCode + "。" + Environment.NewLine + log);
                    return log;
                }
            });
        }

        private static string QuoteArgument(string value)
        {
            return "\"" + (value ?? "").Replace("\"", "\\\"") + "\"";
        }

        private void ChooseImage(bool svgReference)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog { Filter = "图片|*.png;*.jpg;*.jpeg;*.webp;*.bmp|所有文件|*.*" };
            if (dialog.ShowDialog() != true) return;
            UseImageFile(dialog.FileName, svgReference);
        }

        private void UseImageFile(string fileName, bool svgReference)
        {
            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            if (ext != ".png" && ext != ".jpg" && ext != ".jpeg" && ext != ".webp" && ext != ".bmp")
            {
                SetStatus("仅支持 PNG、JPG、WEBP 或 BMP 图片。");
                return;
            }
            var data = "data:" + Mime(fileName) + ";base64," + Convert.ToBase64String(File.ReadAllBytes(fileName));
            if (svgReference)
            {
                _svgReferenceData = NormalizeSvgReference(data);
                _svgReferenceLabel.Text = "参考图: " + Path.GetFileName(fileName); SetPreview(_svgReferencePreview, _svgReferenceData);
                if (_referenceModeBox.SelectedIndex == 0) _referenceModeBox.SelectedIndex = 1;
            }
            else
            {
                _currentImageData = "";
                if (_imageResultPreview != null) _imageResultPreview.Source = null;
                if (_imageResultLabel != null) _imageResultLabel.Text = "";
                _imageSourceData = data; _imageSourceLabel.Text = "来源图片: " + Path.GetFileName(fileName); SetPreview(_imageSourcePreview, data);
            }
            SetStatus(svgReference
                ? "参考图已载入，可直接点击“生成 SVG”；提示词仅用于补充要求。"
                : "图片已载入，可开始处理。");
        }

        private ImageInput SvgReference()
        {
            if (string.IsNullOrWhiteSpace(_svgReferenceData)) return null;
            _config.ReferenceMode = _referenceModeBox.SelectedIndex == 1 ? "copy" : "style";
            return new ImageInput { DataUrl = _svgReferenceData };
        }

            // Keep the reference-image pipeline stable: cap the long edge at 1280px and
        // rasterize onto a white canvas before sending the image to a vision model.
        private static string NormalizeSvgReference(string dataUrl)
        {
            try
            {
                var bytes = DecodeDataUrl(dataUrl);
                using (var input = new MemoryStream(bytes))
                {
                    var source = new BitmapImage();
                    source.BeginInit();
                    source.CacheOption = BitmapCacheOption.OnLoad;
                    source.StreamSource = input;
                    source.EndInit();
                    source.Freeze();
                    if (source.PixelWidth <= 1280 && source.PixelHeight <= 1280 && dataUrl.Length < 1200000)
                        return dataUrl;

                    var scale = Math.Min(1280.0 / source.PixelWidth, 1280.0 / source.PixelHeight);
                    scale = Math.Min(1.0, scale);
                    var width = Math.Max(1, (int)Math.Round(source.PixelWidth * scale));
                    var height = Math.Max(1, (int)Math.Round(source.PixelHeight * scale));
                    var visual = new DrawingVisual();
                    using (var context = visual.RenderOpen())
                    {
                        context.DrawRectangle(Brushes.White, null, new Rect(0, 0, width, height));
                        context.DrawImage(source, new Rect(0, 0, width, height));
                    }
                    var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
                    bitmap.Render(visual);
                    bitmap.Freeze();
                    var encoder = new JpegBitmapEncoder { QualityLevel = 85 };
                    encoder.Frames.Add(BitmapFrame.Create(bitmap));
                    using (var output = new MemoryStream())
                    {
                        encoder.Save(output);
                        return "data:image/jpeg;base64," + Convert.ToBase64String(output.ToArray());
                    }
                }
            }
            catch (Exception ex)
            {
                Log.W("SVG 参考图归一化失败，使用原图: " + ex.Message);
                return dataUrl;
            }
        }

        private GenerationOptions SvgOptions()
        {
            int.TryParse(_widthBox.Text, out var width); int.TryParse(_heightBox.Text, out var height);
            return new GenerationOptions
            {
                Width = width > 0 ? width : 1024, Height = height > 0 ? height : 1024,
                LayerCount = string.IsNullOrWhiteSpace(_layersBox.Text) ? "5" : _layersBox.Text,
                StyleDescription = _styleBox.Text, Palette = _paletteBox.Text,
                NoBackground = _noBackgroundBox.IsChecked == true,
                ReferenceMode = _config.ReferenceMode,
                Temperature = ParseDouble(ChoiceValue(_temperatureBox == null ? "" : _temperatureBox.SelectedItem as string, "0.6"), 0.6), MaxTokens = 16000
            };
        }

        private async Task<string> ToDataUrl(ImageResponse response)
        {
            if (!string.IsNullOrWhiteSpace(response.Base64)) return "data:image/png;base64," + response.Base64;
            if (string.IsNullOrWhiteSpace(response.Url)) throw new InvalidOperationException("接口没有返回图片数据。");
            using (var web = new WebClient())
            {
                var bytes = await web.DownloadDataTaskAsync(response.Url);
                return "data:image/png;base64," + Convert.ToBase64String(bytes);
            }
        }

        private void SetupDropTarget(ScrollViewer target, bool svgWorkspace)
        {
            target.AllowDrop = true;
            target.DragOver += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effects = DragDropEffects.Copy;
                    e.Handled = true;
                }
            };
            target.Drop += (s, e) =>
            {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files == null || files.Length == 0) return;
                HandleDroppedFile(files[0], svgWorkspace);
                e.Handled = true;
            };
        }

        private static void OnFileDragOver(System.Windows.DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void OnSvgDrop(System.Windows.DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;
            HandleDroppedFile(files[0], true);
            e.Handled = true;
        }

        private void OnImageDrop(System.Windows.DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;
            HandleDroppedFile(files[0], false);
            e.Handled = true;
        }

        private void HandleDroppedFile(string path, bool svgWorkspace)
        {
            if (!File.Exists(path)) return;
            var extension = Path.GetExtension(path).ToLowerInvariant();
            try
            {
                if (svgWorkspace && extension == ".svg")
                {
                    var svg = SvgText.Extract(File.ReadAllText(path));
                    if (string.IsNullOrWhiteSpace(svg)) throw new InvalidDataException("拖入的文件不是有效 SVG。");
                    var result = new SvgGenerationResult { Svg = svg, Layers = SvgLayers.Parse(svg) };
                    ApplySvg(result, "外部 SVG");
                    _svgPrompt.Text = "[外部 SVG] " + Path.GetFileName(path);
                    SetStatus("外部 SVG 已载入 ✓ " + Path.GetFileName(path));
                    return;
                }
                if (IsImageExtension(extension))
                {
                    var data = "data:" + Mime(path) + ";base64," + Convert.ToBase64String(File.ReadAllBytes(path));
                    if (svgWorkspace)
                    {
                        _svgReferenceData = NormalizeSvgReference(data);
                        _svgReferenceLabel.Text = "参考图: " + Path.GetFileName(path);
                        SetPreview(_svgReferencePreview, _svgReferenceData);
                        if (_referenceModeBox.SelectedIndex == 0) _referenceModeBox.SelectedIndex = 1;
                    }
                    else
                    {
                        _imageSourceData = data;
                        _imageSourceLabel.Text = "来源图片: " + Path.GetFileName(path);
                        SetPreview(_imageSourcePreview, data);
                    }
                    SetStatus((svgWorkspace ? "参考图" : "来源图片") + "拖入成功 ✓ " + Path.GetFileName(path));
                    return;
                }
                SetStatus("不支持的拖入文件类型: " + extension);
            }
            catch (Exception ex)
            {
                SetStatus("拖入文件处理失败: " + ex.Message);
                Log.W("拖入文件失败: " + ex);
            }
        }

        private static bool IsImageExtension(string extension)
        {
            return extension == ".png" || extension == ".jpg" || extension == ".jpeg" ||
                   extension == ".bmp" || extension == ".webp" || extension == ".gif";
        }

        private string CurrentImageFile()
        {
            if (string.IsNullOrWhiteSpace(_currentImageData)) throw new InvalidOperationException("请先生成图片。");
            return WriteTempFile("image", ".png", DecodeDataUrl(_currentImageData));
        }

        private static byte[] DecodeDataUrl(string data)
        {
            var comma = (data ?? "").IndexOf(',');
            if (comma < 0) throw new InvalidOperationException("图片数据格式错误。");
            return Convert.FromBase64String(data.Substring(comma + 1));
        }

        private string WriteTempFile(string prefix, string ext, byte[] bytes)
        {
            var dir = Path.Combine(_addonDir, "output");
            Directory.CreateDirectory(dir);
            var file = Path.Combine(dir, prefix + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ext);
            File.WriteAllBytes(file, bytes);
            return file;
        }

        private void SaveSvgHistory(SvgGenerationResult result)
        {
            var file = WriteTempFile("svg", ".svg", Encoding.UTF8.GetBytes(result.Svg));
            _history.Insert(0, new SvgHistory { FilePath = file, Prompt = _svgPrompt.Text, Layers = result.Layers.Count });
            while (_history.Count > 200) _history.RemoveAt(_history.Count - 1);
            File.WriteAllLines(_historyPath, _history.Select(x => x.FilePath + "\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes(x.Prompt ?? "")) + "\t" + x.Layers));
            _historyBox.ItemsSource = null; _historyBox.ItemsSource = _history; _historyBox.SelectedIndex = 0;
        }

        private void LoadHistory()
        {
            if (!File.Exists(_historyPath)) return;
            foreach (var line in File.ReadAllLines(_historyPath))
            {
                var parts = line.Split('\t');
                if (parts.Length < 3 || !File.Exists(parts[0])) continue;
                int.TryParse(parts[2], out var layers);
                _history.Add(new SvgHistory { FilePath = parts[0], Prompt = Encoding.UTF8.GetString(Convert.FromBase64String(parts[1])), Layers = layers });
            }
        }

        private void LoadSelectedHistory()
        {
            var history = _historyBox.SelectedItem as SvgHistory;
            if (history == null) return;
            _currentSvg = SvgText.Extract(File.ReadAllText(history.FilePath));
            _svgPrompt.Text = history.Prompt;
            _svgCode.Text = _currentSvg;
            var layers = SvgLayers.Parse(_currentSvg);
            _variants.Add(new SvgVariant { Label = "历史", Svg = _currentSvg, Layers = layers });
            _variantBox.ItemsSource = null; _variantBox.ItemsSource = _variants; _variantBox.SelectedIndex = _variants.Count - 1;
            RenderLayers(layers);
            SetStatus("已载入 SVG 历史。");
        }

        private void FillSettings(ApiProfile p)
        {
            if (p == null) return;
            _profileNameBox.Text = p.Name;
            _apiTypeBox.SelectedItem = p.ApiType;
            _baseUrlBox.Text = p.BaseUrl;
            _apiKeyBox.Text = p.ApiKey;
            _visionBox.IsChecked = p.Vision;
            RefreshSettingsModelSelectors(p);
        }

        private void SaveSettings(int index)
        {
            if (index < 0 || index >= _config.Profiles.Count) return;
            var p = _config.Profiles[index];
            SyncSettingsToProfile(index);
            if (p.Models == null) p.Models = new List<string>();
            foreach (var model in new[] { p.Model, p.ImageModel }.Where(x => !string.IsNullOrWhiteSpace(x) && !p.Models.Contains(x))) p.Models.Add(model);
            _config.ActiveIndex = index;
            SaveConfig();
            RefreshProfiles();
            FillSettings(p);
            SetStatus("模型设置已保存。可在 SVG 模型和生图模型下拉列表中直接选择。");
        }

        private void SyncSettingsToProfile(int index)
        {
            if (index < 0 || index >= _config.Profiles.Count || _profileNameBox == null) return;
            var p = _config.Profiles[index];
            p.Name = _profileNameBox.Text.Trim();
            p.ApiType = (_apiTypeBox.SelectedItem as string) ?? "openai";
            p.BaseUrl = _baseUrlBox.Text.Trim().TrimEnd('/');
            p.Model = SelectedModelText(_settingsModelBox);
            p.ImageModel = SelectedModelText(_settingsImageModelBox);
            p.ApiKey = _apiKeyBox.Text.Trim();
            p.Vision = _visionBox.IsChecked == true;
        }

        private void RefreshSettingsModelSelectors(ApiProfile profile)
        {
            SetModelSelectorItems(_settingsModelBox, profile == null ? null : profile.Models, profile == null ? "" : profile.Model);
            SetModelSelectorItems(_settingsImageModelBox, profile == null ? null : profile.Models, profile == null ? "" : profile.ImageModel);
        }

        private static void SetModelSelectorItems(ComboBox selector, IEnumerable<string> models, string selectedModel)
        {
            if (selector == null) return;
            var items = (models ?? Enumerable.Empty<string>())
                .Where(model => !string.IsNullOrWhiteSpace(model))
                .Select(model => model.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(model => model, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            selector.ItemsSource = null;
            selector.ItemsSource = items;
            selector.SelectedItem = items.FirstOrDefault(model => string.Equals(model, selectedModel, StringComparison.OrdinalIgnoreCase));
            selector.Text = selectedModel ?? "";
        }

        private static string SelectedModelText(ComboBox selector)
        {
            return ((selector == null ? null : selector.SelectedItem) as string ?? (selector == null ? "" : selector.Text) ?? "").Trim();
        }

        private async Task TestApiConnectionAsync(int index)
        {
            if (index < 0 || index >= _config.Profiles.Count) { SetStatus("请先选择模型档案。"); return; }
            try
            {
                SetStatus("正在测试中转连接...");
                using (var client = new ApiClient(_config.Proxy))
                {
                    var models = await client.FetchModelsAsync(_config.Profiles[index], CancellationToken.None);
                    SetStatus("中转连接成功 ✓ " + _config.Profiles[index].BaseUrl + "，可用模型 " + models.Count + " 个。");
                }
            }
            catch (Exception ex)
            {
                SetStatus("中转连接失败: " + ex.Message + " 地址: " + _config.Profiles[index].BaseUrl);
                Log.W("中转连接失败: " + ex);
            }
        }

        private async Task FetchModelsAsync(int index)
        {
            if (index < 0 || index >= _config.Profiles.Count) { SetStatus("请先选择模型档案。"); return; }
            try
            {
                SetStatus("正在获取模型列表...");
                using (var client = new ApiClient(_config.Proxy))
                {
                    var models = await client.FetchModelsAsync(_config.Profiles[index], CancellationToken.None);
                    var profile = _config.Profiles[index];
                    profile.Models = models.ToList();
                    RefreshProfiles();
                    FillSettings(profile);
                    SetStatus("模型列表已更新 ✓ 共 " + models.Count + " 个。现在可点击 SVG 模型或生图模型下拉框选择。");
                }
            }
            catch (Exception ex)
            {
                SetStatus("获取模型列表失败: " + ex.Message + " 地址: " + _config.Profiles[index].BaseUrl);
                Log.W("获取模型列表失败: " + ex);
            }
        }

        private AppConfig LoadConfig()
        {
            try { return AppConfig.Load(_configPath); }
            catch (Exception ex) { Log.W("配置读取失败: " + ex.Message); return new AppConfig(); }
        }

        private void SaveConfig()
        {
            _config.LastPrompt = _svgPrompt == null ? _config.LastPrompt : _svgPrompt.Text;
            _config.SvgWidth = ParseInt(_widthBox, 1024); _config.SvgHeight = ParseInt(_heightBox, 1024);
            _config.LayerCount = _layersBox == null ? _config.LayerCount : _layersBox.Text;
            _config.Palette = _paletteBox == null ? _config.Palette : _paletteBox.Text;
            _config.NoBackground = _noBackgroundBox != null && _noBackgroundBox.IsChecked == true;
            _config.AutoLayer = _autoLayerBox == null ? _config.AutoLayer : _autoLayerBox.IsChecked == true;
            _config.ReverseLayers = _reverseLayerBox == null ? _config.ReverseLayers : _reverseLayerBox.IsChecked == true;
            _config.VariantCount = _variantCountBox == null ? _config.VariantCount : ChoiceValue(_variantCountBox.SelectedItem as string, _config.VariantCount);
            _config.Creativity = _temperatureBox == null ? _config.Creativity : ChoiceValue(_temperatureBox.SelectedItem as string, _config.Creativity);
            _config.StyleIndex = _stylePresetBox == null ? _config.StyleIndex : _stylePresetBox.SelectedIndex;
            _config.ActiveIndex = _profileBox == null ? _config.ActiveIndex : Math.Max(0, _profileBox.SelectedIndex);
            _config.Save(_configPath);
        }

        private static int ParseInt(TextBox box, int fallback) { return box != null && int.TryParse(box.Text, out var value) && value > 0 ? value : fallback; }
        private static double ParseDouble(string value, double fallback) { return double.TryParse(value, out var result) ? result : fallback; }
        private static string ChoiceValue(string text, string fallback)
        {
            var match = System.Text.RegularExpressions.Regex.Match(text ?? "", @"\d+(?:\.\d+)?");
            return match.Success ? match.Value : fallback;
        }
        private static int ChoiceIndex(string text, double first, double step)
        {
            var value = ParseDouble(ChoiceValue(text, first.ToString(System.Globalization.CultureInfo.InvariantCulture)), first);
            var index = (int)Math.Round((value - first) / step);
            return Math.Max(0, Math.Min(2, index));
        }
        private CancellationToken Token { get { return _requestCts == null ? CancellationToken.None : _requestCts.Token; } }
        private void Delta(string text) { Dispatcher.BeginInvoke(new Action(() => { if (_svgCode.Text.Length < 200000) _svgCode.AppendText(text); })); }

        private void BeginBusy(string text)
        {
            _requestCts = new CancellationTokenSource();
            _jobTitle = text;
            _jobOutcome = "运行中";
            _jobStopwatch.Restart();
            _jobTimer.Start();
            RefreshJobStatus();
            SetStatus(text);
            foreach (var button in new[] { _svgGenerateButton, _svgEditButton, _svgRefineButton, _svgImportButton, _imageGenerateButton, _imageImportButton, _imageTraceButton }) if (button != null) button.IsEnabled = false;
        }

        private void EndBusy()
        {
            _jobStopwatch.Stop();
            _jobTimer.Stop();
            if (string.IsNullOrWhiteSpace(_jobOutcome) || _jobOutcome == "运行中") _jobOutcome = "已完成";
            RefreshJobStatus();
            foreach (var button in new[] { _svgGenerateButton, _svgEditButton, _svgRefineButton, _svgImportButton, _imageGenerateButton, _imageImportButton, _imageTraceButton }) if (button != null) button.IsEnabled = true;
            if (_requestCts != null) { _requestCts.Dispose(); _requestCts = null; }
        }

        private void MarkJobOutcome(string outcome)
        {
            _jobOutcome = outcome ?? "";
        }

        private void RefreshJobStatus()
        {
            if (_jobStatus == null) return;
            var elapsed = FormatElapsed(_jobStopwatch.Elapsed);
            if (_jobStopwatch.IsRunning)
                _jobStatus.Text = "作业状态：" + _jobOutcome + " · " + _jobTitle + " · 已用时 " + elapsed;
            else if (!string.IsNullOrWhiteSpace(_jobTitle))
                _jobStatus.Text = "作业状态：" + _jobOutcome + " · " + _jobTitle + " · 总用时 " + elapsed;
            else
                _jobStatus.Text = "作业状态：空闲";
        }

        private static string FormatElapsed(TimeSpan elapsed)
        {
            return elapsed.TotalHours >= 1
                ? string.Format("{0:00}:{1:00}:{2:00}", (int)elapsed.TotalHours, elapsed.Minutes, elapsed.Seconds)
                : string.Format("{0:00}:{1:00}", elapsed.Minutes, elapsed.Seconds);
        }

        private void SetStatus(string text) { _status.Text = text; }

        private CdrApp GetApp(out string via)
        {
            var shared = AppDomain.CurrentDomain.GetData("AIVectorHelper.CorelApp") as CdrApp;
            if (shared != null) { via = "启动器注入"; return shared; }
            if (PluginHost.App != null) { via = "插件入口注入"; return PluginHost.App; }
            via = ""; return null;
        }

        private static bool SetPreview(Image image, string dataUrl)
        {
            try
            {
                var bytes = DecodeDataUrl(dataUrl);
                using (var stream = new MemoryStream(bytes))
                {
                    var bitmap = new BitmapImage(); bitmap.BeginInit(); bitmap.CacheOption = BitmapCacheOption.OnLoad; bitmap.StreamSource = stream; bitmap.EndInit(); bitmap.Freeze();
                    image.Source = bitmap;
                }
                return true;
            }
            catch { image.Source = null; return false; }
        }

        private void UpdateSvgPreview(string svg)
        {
            if (_svgPreview == null || string.IsNullOrWhiteSpace(svg)) return;
            try
            {
                _svgPreview.NavigateToString("<html><body style='margin:0;background:#f2f2f2;display:flex;align-items:center;justify-content:center;'><div style='width:100%;height:100%;'>" + svg + "</div></body></html>");
            }
            catch (Exception ex) { Log.W("SVG 预览失败: " + ex.Message); }
        }

        private static string[] SvgStyles()
        {
            return new[] { "不指定风格", "扁平插画", "极简线稿", "国潮插画", "儿童卡通", "几何图形", "复古海报", "科技图标" };
        }

        private static IReadOnlyList<ImagePreset> ImagePresets()
        {
            return new[]
            {
                new ImagePreset { Name = "不用预设", Prompt = "", Description = "不套用预设，直接在下方描述框输入完整要求。技巧：把 AI 当成你的员工，要求直接说。" },
                new ImagePreset { Name = "糊转清晰", Prompt = "增强画质，严格复刻原图所有元素（人物、文字、产品、背景），不改变风格，只做清晰度处理，无创意变形。", Description = "适合还原照片、素材和大文字。小字很多的图不要用这个，试试高清放大或去掉文字后期重录。" },
                new ImagePreset { Name = "文本还原", Prompt = "修复文字的模糊、不清晰区域，将文字优化为高清、清晰易读的样式；精准识别推理图片中所有文字，确保文字100%无错乱、无遗漏。", Description = "用于还原模糊海报，对文字有一定修复效果。内容太多需分块处理；AI 推理的文字不保证与原文完全一致，请仔细检查。" },
                new ImagePreset { Name = "拆分元素", Prompt = "这是个美陈布置画面，将一块一块的美陈元素拆分并转换为二维平面制作文件，清除透视，垂直对齐正面视角。每一块元素绝对忠实于原始图像的结构、比例及特征，无任何偏差；清除装饰气球。所有元素平铺排列在比例9:21的灰色画布上，确保元素绝对分离，不重叠，依次排列。", Description = "拆分美陈用这个。拆出来不是分层文件，需配合抠图手动处理。元素过多请分多次执行，建议高分辨率。" },
                new ImagePreset { Name = "拆分海报", Prompt = "这是一张海报，把背景图层、文字图层、主体图层、配图图层分开一起平铺排列在比例21:9的灰色画布上，确保每个图层相互分离，不重叠，依次排列。", Description = "拆分海报元素平铺排列。拆出来不是分层文件，需配合抠图手动处理。元素过多请分多次执行。" },
                new ImagePreset { Name = "融合海报", Prompt = "将这个海报不和谐的元素融合，优化细节，提高整体清晰度，并精准识别图片中所有文字，确保文字100%无错乱、无遗漏；pixel-perfect, legible, high-resolution。", Description = "用于融合海报元素使整体和谐。小字建议后期重新录入。" },
                new ImagePreset { Name = "横版转竖", Prompt = "将图中横版海报变为竖版，元素拆分重新排版设计，补齐空白位置让整个画面完整。", Description = "改变海报比例并重新排版。建议先画目标比例矩形，把原海报放中间，两者群组后选中，再执行“从 CDR 选中对象获取”。" },
                new ImagePreset { Name = "竖版转横", Prompt = "将图中竖版海报变为横版，元素拆分重新排版设计，补齐空白位置让整个画面完整。", Description = "先画目标比例矩形与原海报群组后选中，再执行“从 CDR 选中对象获取”。" },
                new ImagePreset { Name = "比例扩图", Prompt = "确保中间画面不变的情况下，以中间画面为准，填充空白区域并与中间画面融合，确保铺满白色画面。不能直接拉伸画面。", Description = "扩大填充图片。建议先画目标比例矩形，原图放中间，群组后选中再执行。" },
                new ImagePreset { Name = "提取主体", Prompt = "提取主体，删除背景，不要透明背景，增强画质，提高清晰度。", Description = "提取主体并删除背景。可在描述框补充想要的背景颜色等细节。" },
                new ImagePreset { Name = "提取背景", Prompt = "移除所有中文、英文、数字等，只保留背景图。", Description = "提取图片中的背景。可在描述框补充要改变的细节。" },
                new ImagePreset { Name = "去掉文字", Prompt = "移除所有文字，保留无字底图。", Description = "去掉图片中的文字。可在描述框补充其他想去掉的元素。" },
                new ImagePreset { Name = "转线稿图", Prompt = "主体转为线稿图，线条干净平滑。", Description = "转为线稿图。可在描述框补充线条颜色和背景颜色。" },
                new ImagePreset { Name = "修老照片", Prompt = "基于提供的参考图像，执行严格的超高分辨率画质增强。核心要求：绝对忠实于原始图像的面部结构、五官比例及人物身份特征，无任何偏差；表情神态、视线方向、人物姿势、相机拍摄角度、画面整体构图及透视关系保持100%一致。服装款式、头发造型、皮肤状态、背景元素完全保留，不得添加、删减或修改任何元素。精准恢复细微层级细节，极致自然写实，拒绝任何风格化、滤镜化处理。色彩与光线严格对齐原图，仅在分辨率、细节清晰度上实现提升。", Description = "修复破损旧照片并上色，多用于人像。" },
                new ImagePreset { Name = "转电子版", Prompt = "这是一张拍摄的海报画面，截取海报部分，清除海报透视，垂直对齐正面视角，移除所有破损划痕和噪点，把它变成电子版海报。全程以还原原始图像+提升画质细节为核心，确保增强后图像与原图高度一致，仅在分辨率、细节清晰度上实现提升。", Description = "把拍摄的海报照片转成电子版。小字很多时建议描述框加“去掉文字”，后期重录。" },
                new ImagePreset { Name = "去布纹理", Prompt = "消除布纹，优化布纹处细节，移除噪点，色块平滑。全程以还原原始图像+提升画质细节为核心，确保增强后图像与原图高度一致。", Description = "提取布纹图案，处理后再描摹矢量效果更好。" },
                new ImagePreset { Name = "产品精修", Prompt = "产品精修，产品置于纯净的纯白背景上。正视图，平视视角，3D渲染，精准还原产品颜色与包装材质（如玻璃的通透、塑料的哑光、金属的光泽）。清除所有指纹、灰尘与瑕疵，让产品看起来崭新，瓶身光影立体感强，提升整体质感。瓶身标签/文字需清晰锐利。光线柔和均匀，符合电商主图标准。", Description = "随手拍的产品图一键精修成电商主图。" },
                new ImagePreset { Name = "菜品补齐", Prompt = "菜品，菜品容器（如盘子、碗、盆等一切装菜的容器）都必须补齐并完整显示。摄影级商业菜品拍摄照片。", Description = "扩充补齐菜品图。建议先画目标比例矩形，原图放中间，群组后选中再执行。" },
                new ImagePreset { Name = "转二维图", Prompt = "转为矢量图风格，严格复刻原图所有元素（人物、文字、产品、背景），不改变风格，无创意变形；去掉所有阴影效果，去掉所有灯光效果。轮廓平滑干净。要尽量减少颜色过渡，临近颜色取中间值填充纯色。", Description = "把立体感图片转为平面二维图，处理后再点“描摹矢量并导入”效果最佳，这是照片转矢量的推荐路线。" },
                new ImagePreset { Name = "高清放大", Prompt = "高清放大：提升分辨率与细节，保持内容与构图完全不变。", Description = "纯放大提清晰度，内容不变。" },
                new ImagePreset { Name = "转矢量图（PowerTRACE 描摹）", Prompt = "", Description = "不调用 AI，优先使用 CDR 内置 PowerTRACE 描摹为可编辑矢量并导入当前文档；适合轮廓清晰的二维图。模糊图先执行“糊转清晰”或“转二维图”。", TraceOnly = true }
            };
        }

        private void UpdateImagePresetDescription()
        {
            var preset = _imagePresetBox == null ? null : _imagePresetBox.SelectedItem as ImagePreset;
            if (_imagePresetDescription != null)
            {
                _imagePresetDescription.Text = preset == null
                    ? ""
                    : (string.IsNullOrWhiteSpace(preset.Prompt)
                        ? "实际提示词：无（仅使用下方补充描述）\n说明：" + preset.Description
                        : "实际提示词：\n" + preset.Prompt + "\n说明：" + preset.Description);
            }
        }

        private void ClearSvgReference()
        {
            _svgReferenceData = "";
            if (_svgReferencePreview != null) _svgReferencePreview.Source = null;
            if (_svgReferenceLabel != null) _svgReferenceLabel.Text = "未选择参考图";
            if (_referenceModeBox != null) _referenceModeBox.SelectedIndex = 0;
            SetStatus("已清除 SVG 参考图。");
        }

        private void ClearImageSource()
        {
            _imageSourceData = "";
            if (_imageSourcePreview != null) _imageSourcePreview.Source = null;
            if (_imageSourceLabel != null) _imageSourceLabel.Text = "未选择来源图片：直接生图；选择后执行图片编辑。";
            SetStatus("已清除图片来源。");
        }

        private void SaveCurrentImage()
        {
            if (string.IsNullOrWhiteSpace(_currentImageData)) { SetStatus("还没有可保存的生成图片。"); return; }
            var dialog = new Microsoft.Win32.SaveFileDialog { Filter = "PNG 图片|*.png|JPEG 图片|*.jpg", FileName = "ai-image.png" };
            if (dialog.ShowDialog() != true) return;
            try
            {
                File.WriteAllBytes(dialog.FileName, DecodeDataUrl(_currentImageData));
                SetStatus("图片已保存: " + dialog.FileName);
            }
            catch (Exception ex) { SetStatus("保存图片失败: " + ex.Message); }
        }

        private void UseImageAsSvgReference()
        {
            if (string.IsNullOrWhiteSpace(_currentImageData)) { SetStatus("还没有可用的生成图片。"); return; }
            _svgReferenceData = _currentImageData;
            _svgReferenceLabel.Text = "参考图：当前 AI 生图结果";
            SetPreview(_svgReferencePreview, _svgReferenceData);
            if (_referenceModeBox != null) _referenceModeBox.SelectedIndex = 2;
            SetStatus("已将当前生图结果设置为 SVG 参考图。");
        }




        private void ChooseSvgFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog { Filter = "SVG 文件|*.svg|所有文件|*.*" };
            if (dialog.ShowDialog() == true) HandleDroppedFile(dialog.FileName, true);
        }

        private void CancelRequest()
        {
            if (_requestCts != null && !_requestCts.IsCancellationRequested)
            {
                _requestCts.Cancel();
                MarkJobOutcome("正在停止");
                RefreshJobStatus();
                SetStatus("已请求停止当前任务。");
            }
        }

        private void SetAllLayers(bool enabled)
        {
            if (_layerPanel == null) return;
            foreach (var check in _layerPanel.Children.OfType<CheckBox>()) check.IsChecked = enabled;
            SetStatus(enabled ? "已选中全部图层。" : "已取消全部图层。");
        }

        private void OpenOutputDirectory()
        {
            var output = Path.Combine(_addonDir, "output");
            Directory.CreateDirectory(output);
            Process.Start(new ProcessStartInfo { FileName = output, UseShellExecute = true });
        }

        private void DeleteSelectedHistory()
        {
            var history = _historyBox == null ? null : _historyBox.SelectedItem as SvgHistory;
            if (history == null) return;
            _history.Remove(history);
            try { if (File.Exists(history.FilePath)) File.Delete(history.FilePath); } catch { }
            File.WriteAllLines(_historyPath, _history.Select(x => x.FilePath + "\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes(x.Prompt ?? "")) + "\t" + x.Layers));
            _historyBox.ItemsSource = null;
            _historyBox.ItemsSource = _history;
            if (_history.Count > 0) _historyBox.SelectedIndex = 0;
            SetStatus("已删除 SVG 历史。");
        }

        private void GrabSelectionFromCdr()
        {
            try
            {
                var app = GetApp(out var via);
                if (app == null) throw new InvalidOperationException("没有取得 CorelDRAW Application。");
                var doc = app.ActiveDocument;
                if (doc == null) throw new InvalidOperationException("CorelDRAW 中没有打开的文档。");

                // 同时兼容 Selection 属性和 Selection() 方法。
                dynamic dynamicDoc = doc;
                dynamic selection = null;
                var count = 0;
                try
                {
                    selection = dynamicDoc.Selection;
                    count = selection == null ? 0 : (int)selection.Shapes.Count;
                }
                catch { }
                if (count == 0)
                {
                    try
                    {
                        selection = dynamicDoc.Selection();
                        count = selection == null ? 0 : (int)selection.Shapes.Count;
                    }
                    catch { }
                }
                if (count == 0) throw new InvalidOperationException("请先在 CDR 中选中要处理的对象。");

                selection.Copy();
                var sourceFile = Path.Combine(_addonDir, "output", "cdr-source_" + DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".png");
                Directory.CreateDirectory(Path.GetDirectoryName(sourceFile));
                ExportClipboardImageWithPowerShell(sourceFile);
                if (!File.Exists(sourceFile) || new FileInfo(sourceFile).Length == 0)
                    throw new InvalidOperationException("未能从剪贴板获取 CDR 选中对象图像，已停止发送，避免误用旧图。");

                var bytes = File.ReadAllBytes(sourceFile);
                var imageInfo = GetPngSize(bytes);
                _imageSourceData = "data:image/png;base64," + Convert.ToBase64String(bytes);
                var previewOk = SetPreview(_imageSourcePreview, _imageSourceData);
                _imageSourceLabel.Text = "来源图片: CDR 选中对象（" + count + " 个）"
                    + (imageInfo == null ? "" : " · " + imageInfo.Item1 + "×" + imageInfo.Item2)
                    + (previewOk ? "" : " · 预览失败");
                Log.W("CDR 选中对象捕获成功: count=" + count + " bytes=" + bytes.Length
                    + " file=" + sourceFile + " preview=" + previewOk);
                SetStatus("已从 CDR 选中对象获取来源图片 ✓ " + via);
            }
            catch (Exception ex)
            {
                SetStatus("获取 CDR 选中对象失败: " + ex.Message);
                Log.W("获取 CDR 选中对象失败: " + ex);
            }
        }

        private static void ExportClipboardImageWithPowerShell(string path)
        {
            var windir = Environment.GetEnvironmentVariable("WINDIR") ?? @"C:\Windows";
            var powershell = Path.Combine(windir, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            if (!File.Exists(powershell)) powershell = "powershell.exe";
            var safePath = (path ?? "").Replace("'", "''");
            var command = "$img=Get-Clipboard -Format Image; if($img){$img.Save('" + safePath + "')}";
            var startInfo = new ProcessStartInfo
            {
                FileName = powershell,
                Arguments = "-NoProfile -NonInteractive -WindowStyle Hidden -Command " + QuoteArgument(command),
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory
            };
            using (var process = new Process { StartInfo = startInfo })
            {
                if (!process.Start()) throw new InvalidOperationException("无法启动 PowerShell。");
                if (!process.WaitForExit(10000))
                {
                    try { process.Kill(); } catch { }
                    throw new InvalidOperationException("PowerShell 获取剪贴板图像超时。");
                }
                if (process.ExitCode != 0)
                    throw new InvalidOperationException("PowerShell 获取剪贴板图像失败，退出码 " + process.ExitCode + "。");
            }
        }

        private static Tuple<int, int> GetPngSize(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 24
                || bytes[0] != 0x89 || bytes[1] != 0x50 || bytes[2] != 0x4E || bytes[3] != 0x47)
                return null;
            var width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
            var height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
            return width > 0 && height > 0 ? Tuple.Create(width, height) : null;
        }

        private static string Mime(string file)
        {
            switch (Path.GetExtension(file).ToLowerInvariant())
            {
                case ".jpg": case ".jpeg": return "image/jpeg";
                case ".webp": return "image/webp";
                case ".bmp": return "image/bmp";
                default: return "image/png";
            }
        }

        private static string[] SvgTemplates()
        {
            return new[] { "提示词模板", "一只微笑柴犬头像徽章 logo, 圆形构图, 橙白深棕三色", "山间日出风景插画：远山、太阳、云和松树", "卡通橘猫坐在月亮上钓星星", "咖啡店招牌：咖啡杯、咖啡豆与麦穗", "国潮锦鲤：锦鲤、祥云、水波和莲花", "科技感盾牌安全图标", "生日贺卡：气球、彩旗、蛋糕和礼物盒", "圣诞主题边框，中间留白" };
        }
    }

    internal static class ButtonExtensions
    {
        public static Button With(this Button button, RoutedEventHandler handler) { button.Click += handler; return button; }
    }
}
