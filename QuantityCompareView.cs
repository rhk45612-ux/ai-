using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using SD = System.Drawing;
using WF = System.Windows.Forms;
using ClosedXML.Excel;

namespace MyApp
{
    public class QuantityCompareView : WF.UserControl
    {
        private const string SHEET_NAME = "MCC 계산서";

        private readonly WF.Button btnSelectOld;
        private readonly WF.Button btnSelectNew;
        private readonly WF.Button btnCompare;
        private readonly WF.Button btnSave;
        private readonly WF.Button btnSearch;
        private readonly WF.TextBox txtOldPath;
        private readonly WF.TextBox txtNewPath;
        private readonly WF.TextBox txtResult;
        private readonly WF.TextBox txtSearch;

        private string oldFile = "";
        private string newFile = "";
        private readonly Dictionary<string, List<(int Row, int Col)>> diffCoordsPerSheet = new();

        private static SD.Font UiFont => new SD.Font("맑은 고딕", 10);

        public QuantityCompareView()
        {
            Dock = WF.DockStyle.Fill;
            BackColor = SD.Color.White;

            // 초기화
            txtOldPath = new WF.TextBox { ReadOnly = true, Width = 500, Font = UiFont };
            txtNewPath = new WF.TextBox { ReadOnly = true, Width = 500, Font = UiFont };
            txtSearch = new WF.TextBox { Width = 200, Font = UiFont };
            txtResult = new WF.TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = WF.ScrollBars.Vertical,
                Dock = WF.DockStyle.Fill,
                Font = UiFont,
                Margin = new WF.Padding(10)
            };

            btnSelectOld = new WF.Button { Text = "기존 엑셀 선택", Font = UiFont, Width = 150, Height = 35 };
            btnSelectNew = new WF.Button { Text = "비교 엑셀 선택", Font = UiFont, Width = 150, Height = 35 };
            btnCompare = new WF.Button { Text = "비교 시작", Font = UiFont, Width = 150, Height = 40 };
            btnSave = new WF.Button { Text = "결과 저장", Font = UiFont, Width = 120, Height = 40 };
            btnSearch = new WF.Button { Text = "검색", Font = UiFont, Width = 60, Height = 30 };

            btnSelectOld.Click += (_, __) => oldFile = SelectExcelFile(txtOldPath);
            btnSelectNew.Click += (_, __) => newFile = SelectExcelFile(txtNewPath);
            btnCompare.Click += (_, __) => CompareExcels();
            btnSave.Click += (_, __) => SaveResult();
            btnSearch.Click += (_, __) => SearchInResult();

            var inputPanel = new WF.TableLayoutPanel
            {
                Dock = WF.DockStyle.Top,
                ColumnCount = 1,
                RowCount = 5,
                AutoSize = true,
                Padding = new WF.Padding(10),
                Margin = new WF.Padding(0)
            };

            inputPanel.Controls.Add(btnSelectOld, 0, 0);
            inputPanel.Controls.Add(txtOldPath, 0, 1);
            inputPanel.Controls.Add(btnSelectNew, 0, 2);
            inputPanel.Controls.Add(txtNewPath, 0, 3);
            inputPanel.Controls.Add(btnCompare, 0, 4);

            btnSelectOld.Margin = btnSelectNew.Margin = new WF.Padding(0, 5, 0, 2);
            txtOldPath.Margin = txtNewPath.Margin = new WF.Padding(0, 0, 0, 10);
            btnCompare.Margin = new WF.Padding(0, 10, 0, 10);

            Controls.Add(inputPanel); // 기존 topPanel 대신 inputPanel을 사용

            // 검색 패널
            var searchPanel = new WF.FlowLayoutPanel
            {
                Dock = WF.DockStyle.Top,
                FlowDirection = WF.FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new WF.Padding(10),
                Margin = new WF.Padding(10)
            };
            searchPanel.Controls.Add(new WF.Label
            {
                Text = "검색어:",
                Font = UiFont,
                AutoSize = true,
                TextAlign = SD.ContentAlignment.MiddleCenter,
                Padding = new WF.Padding(0, 6, 5, 0)
            });
            searchPanel.Controls.Add(txtSearch);
            searchPanel.Controls.Add(btnSearch);

            // 결과 저장 버튼 (하단 중앙)
            var bottomPanel = new WF.Panel
            {
                Dock = WF.DockStyle.Bottom,
                Height = 60
            };
            btnSave.Anchor = WF.AnchorStyles.None;
            btnSave.Location = new SD.Point((bottomPanel.Width - btnSave.Width) / 2, 10);
            btnSave.Anchor = WF.AnchorStyles.Top;
            bottomPanel.Controls.Add(btnSave);

            // 버튼 위치 자동 조정
            bottomPanel.Resize += (_, __) =>
            {
                btnSave.Left = (bottomPanel.Width - btnSave.Width) / 2;
            };

            Controls.Add(txtResult);       // 결과 텍스트박스 (Fill)
            Controls.Add(bottomPanel);     // 저장 버튼 (Bottom)
            Controls.Add(searchPanel);     // 검색 영역
            Controls.Add(inputPanel);        // 상단 버튼 영역
        }

        private static string SelectExcelFile(WF.TextBox targetBox)
        {
            using var ofd = new WF.OpenFileDialog { Filter = "Excel files|*.xlsx", Title = "엑셀 선택" };
            if (ofd.ShowDialog() == WF.DialogResult.OK)
            {
                targetBox.Text = ofd.FileName;
                return ofd.FileName;
            }
            return "";
        }

        private void CompareExcels()
        {
            txtResult.Clear();
            diffCoordsPerSheet.Clear();

            if (!File.Exists(oldFile) || !File.Exists(newFile))
            {
                WF.MessageBox.Show("엑셀 파일을 모두 선택하세요.");
                return;
            }

            try
            {
                var oldBlocks = ExtractBlocks(oldFile);
                var newBlocks = ExtractBlocks(newFile);
                var allKeys = new HashSet<string>(oldBlocks.Keys.Concat(newBlocks.Keys));

                foreach (var key in allKeys.OrderBy(k => k))
                {
                    txtResult.AppendText($"[{key}]\r\n");

                    if (oldBlocks.TryGetValue(key, out var oldData) && newBlocks.TryGetValue(key, out var newData))
                    {
                        var diffs = CompareBlockData(oldData.Data, newData.Data, oldData.StartRow);
                        if (diffs.Count == 0)
                        {
                            txtResult.AppendText(" - 차이점 없음\r\n\r\n");
                        }
                        else
                        {
                            foreach (var d in diffs)
                                txtResult.AppendText($" - {d.Message}\r\n");

                            if (!diffCoordsPerSheet.ContainsKey(SHEET_NAME))
                                diffCoordsPerSheet[SHEET_NAME] = new();

                            diffCoordsPerSheet[SHEET_NAME].AddRange(diffs.Select(d => d.Coord));
                            txtResult.AppendText("\r\n");
                        }
                    }
                    else if (oldBlocks.ContainsKey(key))
                    {
                        txtResult.AppendText(" - 비교 파일에 없음\r\n\r\n");
                    }
                    else
                    {
                        txtResult.AppendText(" - 기존 파일에 없음\r\n\r\n");
                    }
                }
            }
            catch (Exception ex)
            {
                WF.MessageBox.Show($"비교 중 오류 발생:\n{ex.Message}", "오류",
                    WF.MessageBoxButtons.OK, WF.MessageBoxIcon.Error);
            }
        }

        private void SaveResult()
        {
            if (string.IsNullOrEmpty(oldFile) || diffCoordsPerSheet.Count == 0)
            {
                WF.MessageBox.Show("비교 후 저장 가능합니다.");
                return;
            }

            var sfd = new WF.SaveFileDialog { Filter = "Excel files|*.xlsx", Title = "결과 저장", FileName = "비교_결과.xlsx" };
            if (sfd.ShowDialog() != WF.DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook(oldFile);
                var fill = XLColor.LightSkyBlue;

                foreach (var (sheetName, coords) in diffCoordsPerSheet)
                {
                    if (!wb.Worksheets.TryGetWorksheet(sheetName, out var ws)) continue;
                    foreach (var (row, col) in coords)
                        ws.Cell(row, col).Style.Fill.BackgroundColor = fill;
                }

                wb.SaveAs(sfd.FileName);
                WF.MessageBox.Show("엑셀 저장 완료!");
            }
            catch (Exception ex)
            {
                WF.MessageBox.Show("저장 중 오류: " + ex.Message);
            }
        }

        private sealed record Block(int StartRow, int EndRow, List<List<string>> Data);

        private static Dictionary<string, Block> ExtractBlocks(string path)
        {
            var dict = new Dictionary<string, Block>();
            using var wb = new XLWorkbook(path);

            if (!wb.Worksheets.TryGetWorksheet(SHEET_NAME, out var ws))
            {
                WF.MessageBox.Show($"시트 '{SHEET_NAME}'가 존재하지 않습니다.", "오류",
                    WF.MessageBoxButtons.OK, WF.MessageBoxIcon.Error);
                return dict;
            }

            foreach (var range in ws.MergedRanges.ToList()) range.Unmerge();
            foreach (var col in ws.ColumnsUsed()) ws.Column(col.ColumnNumber()).Unhide();
            foreach (var row in ws.RowsUsed()) ws.Row(row.RowNumber()).Unhide();

            int rowCount = ws.LastRowUsed().RowNumber();
            int colCount = ws.LastColumnUsed().ColumnNumber();

            var feederRows = new List<(int Row, int Col)>();
            for (int r = 1; r <= rowCount; r++)
                for (int c = 1; c <= colCount; c++)
                    if (Clean(ws.Cell(r, c).GetString()) == "MCC FEEDER")
                        feederRows.Add((r, c));

            for (int i = 0; i < feederRows.Count; i++)
            {
                int start = feederRows[i].Row;
                int col = feederRows[i].Col;
                int end = (i + 1 < feederRows.Count) ? feederRows[i + 1].Row - 2 : rowCount;

                if (i + 1 == feederRows.Count)
                {
                    for (int r = rowCount; r >= start; r--)
                    {
                        bool hasData = Enumerable.Range(1, colCount)
                            .Any(cc => !string.IsNullOrWhiteSpace(Clean(ws.Cell(r, cc).GetString())));
                        if (hasData) { end = r; break; }
                    }
                }

                string name = Clean(ws.Cell(start + 1, col).GetString());
                if (string.IsNullOrWhiteSpace(name)) name = $"Block_{i + 1}";

                var blockData = new List<List<string>>(capacity: end - start + 1);
                for (int r = start; r <= end; r++)
                {
                    var rowList = new List<string>(capacity: colCount);
                    for (int c = 1; c <= colCount; c++)
                        rowList.Add(Clean(ws.Cell(r, c).GetString()));
                    blockData.Add(rowList);
                }

                dict[name] = new Block(start, end, blockData);
            }

            return dict;
        }

        private static List<(string Message, (int Row, int Col) Coord)> CompareBlockData(
            List<List<string>> a, List<List<string>> b, int startRowInSheet)
        {
            var diffs = new List<(string, (int, int))>();
            int maxRows = Math.Max(a.Count, b.Count);
            int maxCols = Math.Max(a.FirstOrDefault()?.Count ?? 0, b.FirstOrDefault()?.Count ?? 0);

            for (int r = 0; r < maxRows; r++)
            {
                for (int c = 0; c < maxCols; c++)
                {
                    string valA = (r < a.Count && c < a[r].Count) ? a[r][c] : "";
                    string valB = (r < b.Count && c < b[r].Count) ? b[r][c] : "";

                    if (string.IsNullOrWhiteSpace(valA) && string.IsNullOrWhiteSpace(valB))
                        continue;

                    if (!string.Equals(valA, valB, StringComparison.Ordinal))
                    {
                        string colAlpha = ColIndexToName(c);
                        int excelRow = startRowInSheet + r;
                        diffs.Add(($"{excelRow}행 {colAlpha}열: '{valA}' → '{valB}'", (excelRow, c + 1)));
                    }
                }
            }
            return diffs;
        }

        private static string Clean(string s) =>
            Regex.Replace(s ?? "", @"\s+", " ").Trim();

        private static string ColIndexToName(int col)
        {
            col++;
            string result = "";
            while (col > 0)
            {
                col--;
                result = (char)('A' + (col % 26)) + result;
                col /= 26;
            }
            return result;
        }

        private void SearchInResult()
        {
            string keyword = txtSearch.Text.Trim();
            if (string.IsNullOrWhiteSpace(keyword))
            {
                WF.MessageBox.Show("검색어를 입력하세요.");
                return;
            }

            int index = txtResult.Text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
            {
                txtResult.Focus();
                txtResult.SelectionStart = index;
                txtResult.SelectionLength = keyword.Length;
                txtResult.ScrollToCaret();
            }
            else
            {
                WF.MessageBox.Show("검색어를 찾을 수 없습니다.", "검색",
                    WF.MessageBoxButtons.OK, WF.MessageBoxIcon.Information);
            }
        }
    }
}
