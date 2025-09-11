using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using SD = System.Drawing;
using WF = System.Windows.Forms;
using ClosedXML.Excel;

namespace MyApp
{
    public class QuantitySplitView : WF.UserControl
    {
        private const int HeaderRowCount = 3;
        private const int BlockRowSpacing = 2;
        private const string TargetSheetName = "MCC 계산서";

        private static readonly Regex A1RefRegex =
            new Regex(@"(?<![A-Za-z])(\$?[A-Z]{1,3}\$?[0-9]{1,7})",
                      RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private readonly WF.Button btnPick;
        private readonly WF.Button btnSave;
        private readonly WF.FlowLayoutPanel listPanel;

        private readonly Dictionary<string, WF.CheckBox> checkBoxes = new();
        private readonly List<(string Name, int StartRow, int EndRow)> blockList = new();
        private string selectedPath = string.Empty;

        public QuantitySplitView()
        {
            Dock = WF.DockStyle.Fill;
            BackColor = SD.Color.White;

            var btnStack = new WF.FlowLayoutPanel
            {
                FlowDirection = WF.FlowDirection.TopDown,
                AutoSize = true,
                WrapContents = false,
                Margin = new WF.Padding(10, 10, 0, 10),
                Padding = new WF.Padding(0)
            };

            btnPick = new WF.Button { Text = "엑셀 선택", Width = 200, Height = 40, Font = new SD.Font("맑은 고딕", 10) };
            btnSave = new WF.Button { Text = "선택 블록 저장", Width = 200, Height = 40, Font = new SD.Font("맑은 고딕", 10) };
            btnPick.Margin = new WF.Padding(0, 0, 0, 10);
            btnSave.Margin = new WF.Padding(0);
            btnPick.Click += (_, __) => LoadExcel();
            btnSave.Click += (_, __) => SaveSelectedBlocks();

            btnStack.Controls.Add(btnPick);
            btnStack.Controls.Add(btnSave);

            var center = new WF.TableLayoutPanel
            {
                Dock = WF.DockStyle.Top,
                AutoSize = true,
                ColumnCount = 3,
                RowCount = 3,
                Padding = new WF.Padding(0),
                Margin = new WF.Padding(0)
            };
            center.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.Percent, 50));
            center.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.AutoSize));
            center.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.Percent, 50));
            center.RowStyles.Add(new WF.RowStyle(WF.SizeType.Percent, 0));
            center.RowStyles.Add(new WF.RowStyle(WF.SizeType.AutoSize));
            center.RowStyles.Add(new WF.RowStyle(WF.SizeType.Percent, 0));
            center.Controls.Add(btnStack, 0, 1);

            listPanel = new WF.FlowLayoutPanel
            {
                Dock = WF.DockStyle.Fill,
                FlowDirection = WF.FlowDirection.TopDown,
                AutoScroll = true
            };

            var root = new WF.TableLayoutPanel
            {
                Dock = WF.DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            root.RowStyles.Add(new WF.RowStyle(WF.SizeType.AutoSize));
            root.RowStyles.Add(new WF.RowStyle(WF.SizeType.Percent, 100));

            root.Controls.Add(center, 0, 0);
            root.Controls.Add(listPanel, 0, 1);

            Controls.Add(root);
        }

        private void LoadExcel()
        {
            using var ofd = new WF.OpenFileDialog { Filter = "Excel files|*.xlsx", Title = "엑셀 파일 선택" };
            if (ofd.ShowDialog() != WF.DialogResult.OK) return;

            selectedPath = ofd.FileName;
            blockList.Clear();
            checkBoxes.Clear();
            listPanel.Controls.Clear();

            using var workbook = new XLWorkbook(selectedPath);
            if (!workbook.Worksheets.Contains(TargetSheetName))
            {
                WF.MessageBox.Show($"시트 '{TargetSheetName}'를 찾을 수 없습니다.");
                return;
            }
            var ws = workbook.Worksheet(TargetSheetName);

            foreach (var r in ws.MergedRanges.ToList()) r.Unmerge();
            foreach (var c in ws.ColumnsUsed()) c.Unhide();
            foreach (var r in ws.RowsUsed()) r.Unhide();

            int rowCount = ws.LastRowUsed().RowNumber();
            int colCount = ws.LastColumnUsed().ColumnNumber();

            var anchors = new List<(int Row, int Col)>();
            for (int r = 1; r <= rowCount; r++)
                for (int c = 1; c <= colCount; c++)
                    if (Clean(ws.Cell(r, c).GetString()) == "MCC FEEDER")
                        anchors.Add((r, c));

            for (int i = 0; i < anchors.Count; i++)
            {
                int start = anchors[i].Row;
                int col = anchors[i].Col;
                int end = (i + 1 < anchors.Count) ? anchors[i + 1].Row - 2 : rowCount;

                if (i + 1 == anchors.Count)
                {
                    for (int r = rowCount; r >= start; r--)
                    {
                        bool has = Enumerable.Range(1, colCount)
                                  .Any(cc => !string.IsNullOrWhiteSpace(Clean(ws.Cell(r, cc).GetString())));
                        if (has) { end = r; break; }
                    }
                }

                int nameRow = (i + 1 < anchors.Count) ? Math.Max(1, end - 1) : end;
                string name = Clean(ws.Cell(nameRow, col).GetString());
                if (string.IsNullOrWhiteSpace(name)) name = $"Block_{i + 1}";

                blockList.Add((name, start, end));
            }

            foreach (var group in blockList.GroupBy(b => b.Name))
            {
                var cb = new WF.CheckBox
                {
                    Text = $"{group.Key} ({group.Count()}ea)",
                    Width = 400,
                    Font = new SD.Font("맑은 고딕", 11)
                };
                checkBoxes[group.Key] = cb;
                listPanel.Controls.Add(cb);
            }

            WF.MessageBox.Show($"{blockList.Count}개의 블록이 감지되었습니다.");
        }

        private void SaveSelectedBlocks()
        {
            if (string.IsNullOrEmpty(selectedPath))
            {
                WF.MessageBox.Show("먼저 엑셀 파일을 선택하세요.");
                return;
            }

            var selectedNames = new List<string>();
            foreach (var kv in checkBoxes)
                if (kv.Value.Checked) selectedNames.Add(kv.Key);

            if (selectedNames.Count == 0)
            {
                WF.MessageBox.Show("저장할 블록을 선택하세요.");
                return;
            }

            using var sfd = new WF.SaveFileDialog { Filter = "Excel files|*.xlsx", Title = "저장 위치 선택" };
            if (sfd.ShowDialog() != WF.DialogResult.OK) return;

            using var original = new XLWorkbook(selectedPath);
            using var newBook = new XLWorkbook();

            var wsOld = original.Worksheet(TargetSheetName);

            foreach (var sheet in original.Worksheets.Where(s => s.Name != TargetSheetName))
            {
                string newSheetName = sheet.Name;
                int suffix = 1;
                while (newBook.Worksheets.Any(ws => ws.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase)))
                    newSheetName = $"{sheet.Name}_{suffix++}";

                var newSheet = newBook.Worksheets.Add(newSheetName);
                foreach (var cell in sheet.CellsUsed())
                {
                    var dst = newSheet.Cell(cell.Address);
                    dst.Value = cell.Value;
                    try { dst.Style = cell.Style; } catch { }
                }
                newSheet.Hide();
            }

            var wsNew = newBook.Worksheets.Add(TargetSheetName);
            foreach (var col in wsOld.ColumnsUsed())
                wsNew.Column(col.ColumnNumber()).Width = col.Width;

            var refs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int writeRow = 1;

            for (int r = 1; r <= HeaderRowCount; r++)
                writeRow = CopyRow(wsOld, wsNew, r, writeRow, refs);

            foreach (var block in blockList)
            {
                if (!selectedNames.Contains(block.Name)) continue;

                for (int r = block.StartRow; r <= block.EndRow; r++)
                    writeRow = CopyRow(wsOld, wsNew, r, writeRow, refs);

                writeRow += BlockRowSpacing;
            }

            var pending = new Queue<string>(refs);
            var visited = new HashSet<string>(refs, StringComparer.OrdinalIgnoreCase);

            while (pending.Count > 0)
            {
                var addr = pending.Dequeue();
                var src = wsOld.Cell(addr);
                var dst = wsNew.Cell(addr);

                if (dst.IsEmpty())
                {
                    var newly = new List<string>();
                    CopyCellWithReferences(src, dst, newly);
                    foreach (var n in newly)
                        if (visited.Add(n)) pending.Enqueue(n);
                }
            }

            if (wsOld.PageSetup.PrintAreas.Any())
            {
                foreach (var area in wsOld.PageSetup.PrintAreas)
                {
                    try
                    {
                        string ra = $"{area.RangeAddress.FirstAddress}:{area.RangeAddress.LastAddress}";
                        wsNew.PageSetup.PrintAreas.Add(ra);
                    }
                    catch { }
                }
            }

            newBook.SaveAs(sfd.FileName);
            WF.MessageBox.Show("완료되었습니다!");
        }

        private static int CopyRow(IXLWorksheet from, IXLWorksheet to, int fromRow, int toRow, ISet<string> refs)
        {
            to.Row(toRow).Height = from.Row(fromRow).Height;
            int colCount = from.LastColumnUsed()?.ColumnNumber() ?? 0;

            for (int c = 1; c <= colCount; c++)
            {
                var src = from.Cell(fromRow, c);
                var dst = to.Cell(toRow, c);
                CopyCellWithReferences(src, dst, refs);
            }

            return toRow + 1;
        }

        private static void CopyCellWithReferences(IXLCell src, IXLCell dst, ICollection<string> collectRefs)
        {
            try
            {
                var f = src.FormulaA1;
                if (!string.IsNullOrWhiteSpace(f))
                {
                    // 위험한(문맥 의존적인) 수식은 값만 복사
                    bool looksExternalOrStructured =
                        f.IndexOf('!') >= 0 || f.IndexOf('[') >= 0 || f.IndexOf(']') >= 0 ||
                        f.IndexOf('"') >= 0 || f.IndexOf('{') >= 0 || f.IndexOf('}') >= 0;

                    if (looksExternalOrStructured)
                    {
                        dst.Value = src.Value; // 수식 대신 값
                    }
                    else
                    {
                        dst.FormulaA1 = f;
                        if (collectRefs != null)
                        {
                            foreach (Match m in A1RefRegex.Matches(f))
                                collectRefs.Add(m.Value);
                        }
                    }
                }
                else
                {
                    dst.Value = src.Value;
                }
            }
            catch
            {
                dst.Value = src.Value;
            }

            try { dst.Style = src.Style; } catch { }
        }


        private static string Clean(string x) =>
            Regex.Replace(x ?? string.Empty, @"\s+", " ").Trim();
    }
}
