// ExcelUnitSizeView.cs
// 요구사항 반영:
// - 탭: "통과" / "불일치" / "ELB"
// - ELB 열이 있지만 값이 없는 경우 → 일반 블록(통과/불일치 분류)
// - ELB 열에 값이 하나라도 있는 경우 → 블록 전체(계산/스케줄 비교/행 요약 포함)를 ELB 탭으로만 이동
// - 메인/CT 규칙(유니트 개수 기준) — 메인600은 유닛 집계에서 빼지 않으며 별도 면수로 계산/표시
//   · MCCB 4개 이하:      메인 유니트 600 + CT300 + CT300 (폭=1200 → 1면)
//   · MCCB 5~11개:        메인 유니트 600 + CT600 + CT300 (폭=1500 → 1면)
//   · MCCB 12개 이상:     메인 유니트 600 + CT900 + CT300 (폭=1800 → 1면)
// - 인버터는 가산면수로만 합산(예: "→ 인버터 2면")
// - 총면수 = 메인면수 + 일반면수(@1800) + 인버터가산면수
// - +600 관련 로직/표시는 전부 제거됨

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using SD = System.Drawing;
using OfficeOpenXml;

namespace MyApp
{
    public class ExcelUnitSizeView : UserControl
    {
        private Button btnPick;
        private TextBox txtPath;

        private TextBox txtSearch;
        private Button btnFindNext;
        private Button btnClearSearch;

        private TabControl tabs;
        private TabPage tabPass;       // "통과"
        private TabPage tabMismatch;   // "불일치"
        private TabPage tabElb;        // "ELB"

        private TextBox txtPass;
        private TextBox txtMismatch;
        private TextBox txtElb;

        private int searchIndexPass = 0;
        private int searchIndexMismatch = 0;
        private int searchIndexElb = 0;

        public ExcelUnitSizeView()
        {
            Dock = DockStyle.Fill;
            BackColor = SD.Color.White;

            btnPick = new Button
            {
                Text = "엑셀 선택",
                Width = 200,
                Height = 40,
                Font = new SD.Font("맑은 고딕", 10),
                BackColor = SD.Color.White
            };
            btnPick.Click += (_, __) => PickAndAnalyze();

            txtPath = new TextBox
            {
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Font = new SD.Font("맑은 고딕", 10)
            };

            // 검색 영역
            txtSearch = new TextBox { Width = 250, Font = new SD.Font("맑은 고딕", 10) };
            btnFindNext = new Button { Text = "다음 찾기", Width = 90, Height = 28, Font = new SD.Font("맑은 고딕", 9) };
            btnFindNext.Click += (_, __) => SearchInActiveTab(true);
            btnClearSearch = new Button { Text = "초기화", Width = 70, Height = 28, Font = new SD.Font("맑은 고딕", 9) };
            btnClearSearch.Click += (_, __) =>
            {
                txtSearch.Clear();
                txtPass.SelectionLength = 0;
                txtMismatch.SelectionLength = 0;
                txtElb.SelectionLength = 0;
                searchIndexPass = searchIndexMismatch = searchIndexElb = 0;
                txtSearch.Focus();
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
            searchPanel.Controls.Add(txtSearch);
            searchPanel.Controls.Add(btnFindNext);
            searchPanel.Controls.Add(btnClearSearch);

            // 탭/텍스트 박스들
            tabs = new TabControl { Dock = DockStyle.Fill };

            txtPass = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new SD.Font("맑은 고딕", 10)
            };
            txtMismatch = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new SD.Font("맑은 고딕", 10)
            };
            txtElb = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new SD.Font("맑은 고딕", 10)
            };

            tabPass = new TabPage("통과") { Padding = new Padding(6) };
            tabMismatch = new TabPage("불일치") { Padding = new Padding(6) };
            tabElb = new TabPage("ELB") { Padding = new Padding(6) };

            tabPass.Controls.Add(txtPass);
            tabMismatch.Controls.Add(txtMismatch);
            tabElb.Controls.Add(txtElb);

            tabs.TabPages.Add(tabPass);
            tabs.TabPages.Add(tabMismatch);
            tabs.TabPages.Add(tabElb);

            var topPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                RowCount = 2,
                ColumnCount = 1,
                AutoSize = true,
                Padding = new Padding(10)
            };
            topPanel.Controls.Add(btnPick, 0, 0);
            topPanel.Controls.Add(txtPath, 0, 1);

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
            root.Controls.Add(tabs, 0, 2);

            Controls.Add(root);
        }

        private void PickAndAnalyze()
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel files|*.xlsx;*.xls",
                Title = "엑셀 선택"
            };

            if (ofd.ShowDialog() != DialogResult.OK) return;

            txtPath.Text = ofd.FileName;

            var (passText, mismatchText, elbText) = Analyze(ofd.FileName);

            txtPass.Text = string.IsNullOrWhiteSpace(passText) ? "통과 항목 없음" : passText;
            txtMismatch.Text = string.IsNullOrWhiteSpace(mismatchText) ? "불일치 없음" : mismatchText;
            txtElb.Text = string.IsNullOrWhiteSpace(elbText) ? "ELB 없음" : elbText;

            searchIndexPass = searchIndexMismatch = searchIndexElb = 0;
            txtSearch.Focus();
        }

        private static (string pass, string mismatch, string elb) Analyze(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var ws = package.Workbook.Worksheets["MCC 계산서"];
            if (ws == null) return ("", "시트 'MCC 계산서'를 찾을 수 없습니다.", "");
            if (ws.Dimension == null) return ("", "시트에 데이터가 없습니다.", "");

            int rows = ws.Dimension.End.Row;
            int cols = ws.Dimension.End.Column;

            string[,] df = new string[rows, cols];
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                    df[r - 1, c - 1] = Clean(ws.Cells[r, c].Text);

            var blocks = ExtractBlocks(df);
            if (blocks.Count == 0) return ("", "'MCC FEEDER' 앵커를 찾지 못했습니다.", "");

            var passSections = new List<string>();
            var mismatchSections = new List<string>();
            var elbSections = new List<string>();

            for (int bi = 0; bi < blocks.Count; bi++)
            {
                var block = blocks[bi];
                string name = block.name;
                int startRow = block.startRow;
                int endRow = block.endRow;

                var section = new List<string> { $"[{name}]" };

                // 열 인덱스 찾기
                int unitSizeCol = -1, typeCol = -1, myeonCol = -1, elbCol = -1;
                for (int r = startRow; r <= endRow; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string val = Normalize(df[r, c]);
                        if (val == "UNIT SIZE") unitSizeCol = c;
                        if (val == "TYPE") typeCol = c;
                        if (val.Replace(" ", "") == "면수") myeonCol = c;
                        if (val.Replace(" ", "") == "ELB(AF/AT)") elbCol = c;
                    }
                }

                // 스케줄 면수 추출
                int? scheduleMyeon = null;
                if (myeonCol != -1)
                {
                    var nums = new List<int>();
                    for (int r = startRow + 1; r <= endRow; r++)
                    {
                        string s = Clean(df[r, myeonCol]);
                        var ms = Regex.Matches(s, @"\d+");
                        for (int i = 0; i < ms.Count; i++)
                        {
                            if (int.TryParse(ms[i].Value, out int n)) nums.Add(n);
                        }
                    }
                    if (nums.Count > 0) scheduleMyeon = nums[nums.Count - 1];
                    section.Add($"→ 스케줄 표 면수: {(scheduleMyeon?.ToString() ?? "없음")}");
                }
                else
                {
                    section.Add("→ 스케줄 표 면수: 없음");
                }

                // 일반 유닛 카운트/폭
                var normalCounts = new Dictionary<string, int> { ["600"] = 0, ["800"] = 0, ["900"] = 0, ["1200"] = 0 };
                int normalWidthSum = 0;

                // 인버터 가산면수
                int inverterAddFaces = 0;

                if (unitSizeCol != -1)
                {
                    for (int r = startRow + 1; r <= endRow; r++)
                    {
                        string raw = Clean(df[r, unitSizeCol]);
                        if (string.IsNullOrWhiteSpace(raw)) continue;

                        string up = raw.ToUpperInvariant();

                        // 인버터: 가산면만 계산 (일반 유닛 폭/개수에는 미반영)
                        if (up.Contains("W:800"))
                        {
                            inverterAddFaces += 2; // 800 → +2면
                            continue;
                        }
                        if (up.Contains("W:600"))
                        {
                            if (typeCol != -1)
                            {
                                string t = Clean(df[r, typeCol]);
                                if (t == "RI3S12O7L1G-1") inverterAddFaces += 1;
                                else if (t == "RI3S12O7L1G-2") inverterAddFaces += 2;
                            }
                            continue;
                        }

                        // 일반 유닛 집계
                        if (normalCounts.TryGetValue(raw, out int cnt))
                        {
                            normalCounts[raw] = cnt + 1;
                            if (raw == "600") normalWidthSum += 600;
                            else if (raw == "800") normalWidthSum += 800;
                            else if (raw == "900") normalWidthSum += 900;
                            else if (raw == "1200") normalWidthSum += 1200;
                        }
                    }
                }

                int mccbCount = 0;
                foreach (var kv in normalCounts) mccbCount += kv.Value;

                int mainCtWidth = 0;
                string mainLabel = null;
                bool mainAbsorbedUnit600 = false;

                if (mccbCount > 0)
                {
                    if (mccbCount <= 4)
                    {
                        mainCtWidth = 600 + 300 + 300; // 1200 → 1면
                        mainLabel = "MCCB 4개 이하: 메인 유니트 600 + CT300 + CT300";

                        // ✅ 4개 이하일 때 일반 600 한 개를 메인 첫 면에 흡수
                        if (normalCounts.TryGetValue("600", out int c600) && c600 > 0)
                        {
                            normalCounts["600"] = c600 - 1;
                            normalWidthSum -= 600;
                            mainCtWidth += 600; // 1200 + 600 = 1800 → 1면 유지
                            mainAbsorbedUnit600 = true;
                        }
                    }
                    else if (mccbCount <= 11)
                    {
                        mainCtWidth = 600 + 600 + 300; // 1500 → 1면
                        mainLabel = "MCCB 5개 이상 ~ 11개 이하: 메인 유니트 600 + CT600 + CT300";
                    }
                    else
                    {
                        mainCtWidth = 600 + 900 + 300; // 1800 → 1면
                        mainLabel = "MCCB 12개 이상 ~: 메인 유니트 600 + CT900 + CT300";
                    }
                }
                else if (inverterAddFaces > 0)
                {
                    // 인버터만 있는 경우에도 첫 면은 반드시 존재
                    mainCtWidth = 600 + 300;          // 900 → 1면
                    mainLabel = "인버터 전용: 메인 유니트 600 + CT300";
                }

                int mainFaces = mainCtWidth > 0 ? (int)Math.Ceiling(mainCtWidth / 1800.0) : 0;

                if (mainFaces > 0)
                {
                    section.Add(mainAbsorbedUnit600
                        ? $"→ 메인 ({mainLabel} + 600:1EA 포함), {mainFaces}면"
                        : $"→ 메인 ({mainLabel}), {mainFaces}면");
                }
                // 인버터 면수 라인 (있을 때만)
                if (inverterAddFaces > 0)
                    section.Add($"→ 인버터 {inverterAddFaces}면");

                // 일반 유닛 면수(@1800) — +600 관련 로직 없음
                int normalFaces = normalWidthSum > 0 ? (int)Math.Ceiling(normalWidthSum / 1800.0) : 0;

                // 총 면수 = 메인면수 + 일반면수 + 인버터 가산면수
                int totalMyeon = mainFaces + normalFaces + inverterAddFaces;

                // 유닛 요약(간결)
                var parts = new List<string>();
                foreach (var kv in normalCounts)
                {
                    if (kv.Value > 0) parts.Add($"{kv.Key}: {kv.Value}ea");
                }
                if (normalFaces > 0)
                    parts.Add($"{normalFaces}면(@1800)");

                if (parts.Count > 0)
                    section.Add($"→ {string.Join(", ", parts)}");

                // 총면수 라인
                section.Add($"→ 총 계산 면수: {totalMyeon}면");

                // ELB 존재 여부 (ELB 열에 값이 하나라도 있으면 ELB 탭으로만)
                bool hasElbData = false;
                if (elbCol != -1)
                {
                    for (int r = startRow + 1; r <= endRow; r++)
                    {
                        string v = df[r, elbCol];
                        if (!string.IsNullOrWhiteSpace(v)) { hasElbData = true; break; }
                    }
                }

                // 분류
                if (hasElbData)
                {
                    elbSections.Add(string.Join(Environment.NewLine, section));
                    elbSections.Add("");
                }
                else
                {
                    if (scheduleMyeon.HasValue)
                    {
                        if (scheduleMyeon.Value == totalMyeon)
                        {
                            passSections.Add(string.Join(Environment.NewLine, section));
                            passSections.Add("");
                        }
                        else
                        {
                            mismatchSections.Add(string.Join(Environment.NewLine, section));
                            mismatchSections.Add("");
                        }
                    }
                    else
                    {
                        passSections.Add(string.Join(Environment.NewLine, section));
                        passSections.Add("");
                    }
                }
            }

            return (
                string.Join(Environment.NewLine, passSections),
                string.Join(Environment.NewLine, mismatchSections),
                string.Join(Environment.NewLine, elbSections)
            );
        }

        private static string Clean(string input) =>
            string.IsNullOrWhiteSpace(input) ? "" : Regex.Replace(input, @"\s+", " ").Trim();

        private static string Normalize(string input) => Clean(input).ToUpperInvariant();

        /// <summary>MCC FEEDER 기준으로 블록 범위를 찾는다.</summary>
        private static List<(string name, int startRow, int endRow)> ExtractBlocks(string[,] df)
        {
            var blocks = new List<(string name, int startRow, int endRow)>();
            int rows = df.GetLength(0);
            int cols = df.GetLength(1);

            var anchors = new List<(int row, int col, string name)>();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (Normalize(df[r, c]) == "MCC FEEDER")
                    {
                        string nm = (r + 1 < rows) ? Clean(df[r + 1, c]) : $"Block_{anchors.Count + 1}";
                        if (string.IsNullOrWhiteSpace(nm)) nm = $"Block_{anchors.Count + 1}";
                        anchors.Add((r, c, nm));
                    }
                }
            }

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
                        if (hasData) { end = r; break; }
                    }
                }

                blocks.Add((anchors[i].name, start, end));
            }

            return blocks;
        }

        private void SearchInActiveTab(bool next)
        {
            string keyword = txtSearch.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                MessageBox.Show("검색어를 입력하세요.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtSearch.Focus();
                return;
            }

            TextBox target;
            ref int startIndexRef = ref searchIndexPass;

            if (tabs.SelectedTab == tabMismatch)
            {
                target = txtMismatch;
                startIndexRef = ref searchIndexMismatch;
            }
            else if (tabs.SelectedTab == tabElb)
            {
                target = txtElb;
                startIndexRef = ref searchIndexElb;
            }
            else
            {
                target = txtPass;
                startIndexRef = ref searchIndexPass;
            }

            string text = target.Text ?? "";
            if (text.Length == 0)
            {
                MessageBox.Show("검색할 내용이 없습니다.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int idx = text.IndexOf(keyword, startIndexRef, StringComparison.OrdinalIgnoreCase);
            if (idx < 0 && startIndexRef > 0) idx = text.IndexOf(keyword, 0, StringComparison.OrdinalIgnoreCase);

            if (idx >= 0)
            {
                target.Focus();
                target.SelectionStart = idx;
                target.SelectionLength = keyword.Length;
                target.ScrollToCaret();
                startIndexRef = idx + keyword.Length;
            }
            else
            {
                MessageBox.Show("더 이상 검색 결과가 없습니다.", "검색", MessageBoxButtons.OK, MessageBoxIcon.Information);
                startIndexRef = 0;
            }
        }
    }
}
