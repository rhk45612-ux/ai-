using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;
using SD = System.Drawing;

namespace MyApp.Views
{
    public class ExcelUnitSizeView : UserControl
    {
        private readonly Button _btnPick;
        private readonly TextBox _txtPath;

        private readonly TextBox _txtSearch;
        private readonly Button _btnFindNext;
        private readonly Button _btnClearSearch;

        private readonly TabControl _tabs;
        private readonly TabPage _tabPass;
        private readonly TabPage _tabMismatch;
        private readonly TabPage _tabElb;

        private readonly TextBox _txtPass;
        private readonly TextBox _txtMismatch;
        private readonly TextBox _txtElb;

        private readonly Dictionary<TabPage, int> _searchPositions;

        public ExcelUnitSizeView()
        {
            Dock = DockStyle.Fill;
            BackColor = SD.Color.White;

            _btnPick = new Button
            {
                Text = "엑셀 선택",
                Width = 200,
                Height = 40,
                Font = new SD.Font("맑은 고딕", 10),
                BackColor = SD.Color.White
            };
            _btnPick.Click += (_, __) => PickAndAnalyze();

            _txtPath = new TextBox
            {
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Font = new SD.Font("맑은 고딕", 10)
            };

            _txtSearch = new TextBox { Width = 250, Font = new SD.Font("맑은 고딕", 10) };
            _btnFindNext = new Button { Text = "다음 찾기", Width = 90, Height = 28, Font = new SD.Font("맑은 고딕", 9) };
            _btnFindNext.Click += OnFindNext;
            _btnClearSearch = new Button { Text = "초기화", Width = 70, Height = 28, Font = new SD.Font("맑은 고딕", 9) };
            _btnClearSearch.Click += (_, __) =>
            {
                ResetSearchState(clearKeyword: true);
                _txtSearch.Focus();
            };

            var searchPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new Padding(10, 5, 10, 5)
            };
            searchPanel.Controls.Add(new Label
            {
                Text = "검색어:",
                AutoSize = true,
                Font = new SD.Font("맑은 고딕", 10),
                Padding = new Padding(0, 6, 6, 0)
            });
            searchPanel.Controls.Add(_txtSearch);
            searchPanel.Controls.Add(_btnFindNext);
            searchPanel.Controls.Add(_btnClearSearch);

            _tabs = new TabControl { Dock = DockStyle.Fill };

            _txtPass = CreateResultTextBox();
            _txtMismatch = CreateResultTextBox();
            _txtElb = CreateResultTextBox();

            _tabPass = new TabPage("통과") { Padding = new Padding(6) };
            _tabMismatch = new TabPage("불일치") { Padding = new Padding(6) };
            _tabElb = new TabPage("ELB") { Padding = new Padding(6) };

            _tabPass.Controls.Add(_txtPass);
            _tabMismatch.Controls.Add(_txtMismatch);
            _tabElb.Controls.Add(_txtElb);

            _tabs.TabPages.Add(_tabPass);
            _tabs.TabPages.Add(_tabMismatch);
            _tabs.TabPages.Add(_tabElb);

            var topPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                RowCount = 2,
                ColumnCount = 1,
                AutoSize = true,
                Padding = new Padding(10)
            };
            topPanel.Controls.Add(_btnPick, 0, 0);
            topPanel.Controls.Add(_txtPath, 0, 1);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(topPanel, 0, 0);
            root.Controls.Add(searchPanel, 0, 1);
            root.Controls.Add(_tabs, 0, 2);

            Controls.Add(root);

            _searchPositions = new Dictionary<TabPage, int>
            {
                [_tabPass] = 0,
                [_tabMismatch] = 0,
                [_tabElb] = 0
            };
        }

        private static TextBox CreateResultTextBox() => new()
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill,
            Font = new SD.Font("맑은 고딕", 10)
        };

        private void OnFindNext(object? sender, EventArgs e) => ExecuteSearch();

        private void ExecuteSearch()
        {
            string keyword = _txtSearch.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                MessageBox.Show("검색어를 입력하세요.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _txtSearch.Focus();
                return;
            }

            var activePage = _tabs.SelectedTab ?? _tabPass;
            var target = GetTargetTextBox(activePage);

            if (target == null || string.IsNullOrEmpty(target.Text))
            {
                MessageBox.Show("검색할 내용이 없습니다.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!_searchPositions.TryGetValue(activePage, out int startIndex))
                startIndex = 0;

            int idx = target.Text.IndexOf(keyword, startIndex, StringComparison.OrdinalIgnoreCase);
            if (idx < 0 && startIndex > 0)
                idx = target.Text.IndexOf(keyword, 0, StringComparison.OrdinalIgnoreCase);

            if (idx >= 0)
            {
                target.Focus();
                target.SelectionStart = idx;
                target.SelectionLength = keyword.Length;
                target.ScrollToCaret();
                _searchPositions[activePage] = idx + keyword.Length;
            }
            else
            {
                MessageBox.Show("더 이상 검색 결과가 없습니다.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                _searchPositions[activePage] = 0;
            }
        }

        private TextBox? GetTargetTextBox(TabPage page)
        {
            if (page == _tabMismatch) return _txtMismatch;
            if (page == _tabElb) return _txtElb;
            return _txtPass;
        }

        private void ResetSearchState(bool clearKeyword)
        {
            if (clearKeyword)
                _txtSearch.Clear();

            _txtPass.SelectionLength = 0;
            _txtMismatch.SelectionLength = 0;
            _txtElb.SelectionLength = 0;

            _searchPositions[_tabPass] = 0;
            _searchPositions[_tabMismatch] = 0;
            _searchPositions[_tabElb] = 0;
        }

        private void PickAndAnalyze()
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel files|*.xlsx;*.xls",
                Title = "엑셀 선택"
            };

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            _txtPath.Text = ofd.FileName;

            var (passText, mismatchText, elbText) = ExcelUnitSizeAnalyzer.Analyze(ofd.FileName);

            _txtPass.Text = string.IsNullOrWhiteSpace(passText) ? "통과 항목 없음" : passText;
            _txtMismatch.Text = string.IsNullOrWhiteSpace(mismatchText) ? "불일치 없음" : mismatchText;
            _txtElb.Text = string.IsNullOrWhiteSpace(elbText) ? "ELB 없음" : elbText;

            ResetSearchState(clearKeyword: false);
            _txtSearch.Focus();
        }

        private static class ExcelUnitSizeAnalyzer
        {
            private static readonly Regex WsRegex = new("\\s+", RegexOptions.Compiled);

            public static (string pass, string mismatch, string elb) Analyze(string filePath)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets["MCC 계산서"];
                if (worksheet == null)
                    return ("", "시트 'MCC 계산서'를 찾을 수 없습니다.", "");
                if (worksheet.Dimension == null)
                    return ("", "시트에 데이터가 없습니다.", "");

                int rows = worksheet.Dimension.End.Row;
                int cols = worksheet.Dimension.End.Column;

                var df = new string[rows, cols];
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                        df[r - 1, c - 1] = Clean(worksheet.Cells[r, c].Text);
                }

                var blocks = ExtractBlocks(df);
                if (blocks.Count == 0)
                    return ("", "'MCC FEEDER' 앵커를 찾지 못했습니다.", "");

                var passSections = new List<string>();
                var mismatchSections = new List<string>();
                var elbSections = new List<string>();

                var processor = new ExcelUnitBlockProcessor(df, cols);
                foreach (var block in blocks)
                {
                    var result = processor.Process(block);

                    var targetList = result.Category switch
                    {
                        BlockCategory.Pass => passSections,
                        BlockCategory.Mismatch => mismatchSections,
                        BlockCategory.Elb => elbSections,
                        _ => passSections
                    };

                    targetList.Add(string.Join(Environment.NewLine, result.SectionLines));
                    targetList.Add(string.Empty);
                }

                string JoinSections(List<string> sections) => string.Join(Environment.NewLine, sections);

                return (JoinSections(passSections), JoinSections(mismatchSections), JoinSections(elbSections));
            }

            private static string Clean(string input) =>
                string.IsNullOrWhiteSpace(input) ? string.Empty : WsRegex.Replace(input, " ").Trim();

            private static string Normalize(string input) => Clean(input).ToUpperInvariant();

            private static List<(string name, int startRow, int endRow)> ExtractBlocks(string[,] df)
            {
                if (df == null) throw new ArgumentNullException(nameof(df));

                var blocks = new List<(string name, int startRow, int endRow)>();
                int rows = df.GetLength(0);
                int cols = df.GetLength(1);

                if (rows == 0 || cols == 0)
                    return blocks;

                var anchors = new List<(int row, int col, string name)>();

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        if (Normalize(df[r, c]) == "MCC FEEDER")
                        {
                            string name = (r + 1 < rows) ? Clean(df[r + 1, c]) : $"Block_{anchors.Count + 1}";
                            if (string.IsNullOrEmpty(name))
                                name = $"Block_{anchors.Count + 1}";
                            anchors.Add((r, c, name));
                        }
                    }
                }

                if (anchors.Count == 0)
                    return blocks;

                for (int i = 0; i < anchors.Count; i++)
                {
                    int start = anchors[i].row;
                    int end = (i + 1 < anchors.Count) ? anchors[i + 1].row - 2 : rows - 1;

                    if (i + 1 == anchors.Count)
                    {
                        for (int r = rows - 1; r >= start; r--)
                        {
                            bool hasData = false;
                            for (int c = 0; c < cols; c++)
                            {
                                if (!string.IsNullOrWhiteSpace(Clean(df[r, c])))
                                {
                                    hasData = true;
                                    break;
                                }
                            }
                            if (hasData)
                            {
                                end = r;
                                break;
                            }
                        }
                    }

                    if (end < start)
                        end = start;

                    blocks.Add((anchors[i].name, start, end));
                }

                return blocks;
            }

            private enum BlockCategory
            {
                Pass,
                Mismatch,
                Elb
            }

            private sealed record BlockResult(List<string> SectionLines, BlockCategory Category);

            private sealed class ExcelUnitBlockProcessor
            {
                private readonly string[,] _df;
                private readonly int _cols;

                public ExcelUnitBlockProcessor(string[,] df, int cols)
                {
                    _df = df ?? throw new ArgumentNullException(nameof(df));
                    _cols = cols;
                }

                public BlockResult Process((string name, int startRow, int endRow) block)
                {
                    var (name, startRow, endRow) = block;
                    var section = new List<string> { $"[{name}]" };

                    var columns = FindColumns(startRow, endRow);
                    int? scheduleMyeon = ExtractScheduleMyeon(startRow, endRow, columns.MyeonCol);

                    section.Add(columns.MyeonCol != -1
                        ? $"→ 스케줄 표 면수: {(scheduleMyeon?.ToString() ?? "없음")}"
                        : "→ 스케줄 표 면수: 없음");

                    var units = CountUnits(startRow, endRow, columns);
                    var main = ApplyMainRule(units);

                    if (main.Faces > 0 && !string.IsNullOrEmpty(main.Label))
                    {
                        string mainText = main.AbsorbedUnit600
                            ? $"→ 메인 ({main.Label} + 600:1EA 포함), {main.Faces}면"
                            : $"→ 메인 ({main.Label}), {main.Faces}면";
                        section.Add(mainText);
                    }

                    if (units.InverterAddFaces > 0)
                        section.Add($"→ 인버터 {units.InverterAddFaces}면");

                    int normalFaces = units.NormalWidthSum > 0
                        ? (int)Math.Ceiling(units.NormalWidthSum / 1800.0)
                        : 0;

                    var parts = new List<string>();
                    foreach (var kv in units.NormalCounts)
                    {
                        if (kv.Value > 0)
                            parts.Add($"{kv.Key}: {kv.Value}ea");
                    }
                    if (normalFaces > 0)
                        parts.Add($"{normalFaces}면(@1800)");

                    if (parts.Count > 0)
                        section.Add($"→ {string.Join(", ", parts)}");

                    int totalMyeon = main.Faces + normalFaces + units.InverterAddFaces;
                    section.Add($"→ 총 계산 면수: {totalMyeon}면");

                    bool hasElbData = HasElbData(startRow, endRow, columns.ElbCol);
                    var category = DetermineCategory(scheduleMyeon, totalMyeon, hasElbData);

                    return new BlockResult(section, category);
                }

                private ColumnMap FindColumns(int startRow, int endRow)
                {
                    var map = new ColumnMap();
                    for (int r = startRow; r <= endRow; r++)
                    {
                        for (int c = 0; c < _cols; c++)
                        {
                            string value = Normalize(_df[r, c]);
                            if (value == "UNIT SIZE") map.UnitSizeCol = c;
                            if (value == "TYPE") map.TypeCol = c;
                            if (value.Replace(" ", string.Empty) == "면수") map.MyeonCol = c;
                            if (value.Replace(" ", string.Empty) == "ELB(AF/AT)") map.ElbCol = c;
                        }
                    }
                    return map;
                }

                private int? ExtractScheduleMyeon(int startRow, int endRow, int myeonCol)
                {
                    if (myeonCol == -1)
                        return null;

                    var numbers = new List<int>();
                    for (int r = startRow + 1; r <= endRow; r++)
                    {
                        string content = Clean(_df[r, myeonCol]);
                        foreach (Match match in Regex.Matches(content, @"\d+"))
                        {
                            if (int.TryParse(match.Value, out int n))
                                numbers.Add(n);
                        }
                    }

                    return numbers.Count > 0 ? numbers[numbers.Count - 1] : (int?)null;
                }

                private UnitAggregation CountUnits(int startRow, int endRow, ColumnMap columns)
                {
                    var agg = new UnitAggregation();
                    if (columns.UnitSizeCol == -1)
                        return agg;

                    for (int r = startRow + 1; r <= endRow; r++)
                    {
                        string raw = Clean(_df[r, columns.UnitSizeCol]);
                        if (string.IsNullOrEmpty(raw))
                            continue;

                        string upper = raw.ToUpperInvariant();
                        if (upper.Contains("W:800"))
                        {
                            agg.InverterAddFaces += 2;
                            continue;
                        }
                        if (upper.Contains("W:600"))
                        {
                            if (columns.TypeCol != -1)
                            {
                                string type = Clean(_df[r, columns.TypeCol]);
                                if (type == "RI3S12O7L1G-1") agg.InverterAddFaces += 1;
                                else if (type == "RI3S12O7L1G-2") agg.InverterAddFaces += 2;
                            }
                            continue;
                        }

                        if (agg.NormalCounts.TryGetValue(raw, out int count))
                        {
                            agg.NormalCounts[raw] = count + 1;
                            agg.NormalWidthSum += raw switch
                            {
                                "600" => 600,
                                "800" => 800,
                                "900" => 900,
                                "1200" => 1200,
                                _ => 0
                            };
                        }
                    }

                    return agg;
                }

                private MainResult ApplyMainRule(UnitAggregation agg)
                {
                    int mainCtWidth = 0;
                    string? label = null;
                    bool absorbed600 = false;

                    int mccbCount = agg.GetMccbCount();

                    if (mccbCount > 0)
                    {
                        if (mccbCount <= 4)
                        {
                            mainCtWidth = 600 + 300 + 300;
                            label = "MCCB 4개 이하: 메인 유니트 600 + CT300 + CT300";

                            if (agg.RemoveNormalUnit("600", 600))
                            {
                                mainCtWidth += 600;
                                absorbed600 = true;
                            }
                        }
                        else if (mccbCount <= 11)
                        {
                            mainCtWidth = 600 + 600 + 300;
                            label = "MCCB 5개 이상 ~ 11개 이하: 메인 유니트 600 + CT600 + CT300";
                        }
                        else
                        {
                            mainCtWidth = 600 + 900 + 300;
                            label = "MCCB 12개 이상 ~: 메인 유니트 600 + CT900 + CT300";
                        }
                    }
                    else if (agg.InverterAddFaces > 0)
                    {
                        mainCtWidth = 600 + 300;
                        label = "인버터 전용: 메인 유니트 600 + CT300";
                    }

                    int faces = mainCtWidth > 0 ? (int)Math.Ceiling(mainCtWidth / 1800.0) : 0;
                    return new MainResult(faces, label ?? string.Empty, absorbed600);
                }

                private bool HasElbData(int startRow, int endRow, int elbCol)
                {
                    if (elbCol == -1)
                        return false;

                    for (int r = startRow + 1; r <= endRow; r++)
                    {
                        if (!string.IsNullOrWhiteSpace(_df[r, elbCol]))
                            return true;
                    }

                    return false;
                }

                private static BlockCategory DetermineCategory(int? scheduleMyeon, int totalMyeon, bool hasElb)
                {
                    if (hasElb)
                        return BlockCategory.Elb;
                    if (scheduleMyeon.HasValue && scheduleMyeon.Value != totalMyeon)
                        return BlockCategory.Mismatch;
                    return BlockCategory.Pass;
                }

                private sealed class ColumnMap
                {
                    public int UnitSizeCol = -1;
                    public int TypeCol = -1;
                    public int MyeonCol = -1;
                    public int ElbCol = -1;
                }

                private sealed class UnitAggregation
                {
                    public Dictionary<string, int> NormalCounts { get; } = new()
                    {
                        ["600"] = 0,
                        ["800"] = 0,
                        ["900"] = 0,
                        ["1200"] = 0
                    };

                    public int NormalWidthSum { get; set; }
                    public int InverterAddFaces { get; set; }

                    public int GetMccbCount()
                    {
                        int total = 0;
                        foreach (var kv in NormalCounts)
                            total += kv.Value;
                        return total;
                    }

                    public bool RemoveNormalUnit(string key, int width)
                    {
                        if (NormalCounts.TryGetValue(key, out int count) && count > 0)
                        {
                            NormalCounts[key] = count - 1;
                            NormalWidthSum -= width;
                            return true;
                        }
                        return false;
                    }
                }

                private sealed record MainResult(int Faces, string Label, bool AbsorbedUnit600);
            }
        }
    }
}
