using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SD = System.Drawing;
using WF = System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MyApp
{
    /// <summary>
    /// 간단한 워드 문서 비교 뷰.
    /// 두 개의 DOCX 파일을 선택해 문단 단위로 비교하고
    /// 차이를 색상으로 표시합니다.
    /// </summary>
    public class WordComparatorView : WF.UserControl
    {
        private readonly WF.TextBox txtBasePath;
        private readonly WF.TextBox txtComparePath;
        private readonly WF.RichTextBox txtResult;
        private string baseFile = string.Empty;
        private string compareFile = string.Empty;

        public WordComparatorView()
        {
            Dock = WF.DockStyle.Fill;
            BackColor = SD.Color.White;
            var font = new SD.Font("맑은 고딕", 10);

            var btnSelectBase = new WF.Button { Text = "원본 선택", Width = 120, Height = 30, Font = font };
            var btnSelectCompare = new WF.Button { Text = "비교 선택", Width = 120, Height = 30, Font = font };
            var btnCompare = new WF.Button { Text = "비교", Width = 120, Height = 35, Font = font };
            var btnSave = new WF.Button { Text = "결과 저장", Width = 120, Height = 35, Font = font };

            txtBasePath = new WF.TextBox { ReadOnly = true, Width = 500, Font = font };
            txtComparePath = new WF.TextBox { ReadOnly = true, Width = 500, Font = font };
            txtResult = new WF.RichTextBox { Dock = WF.DockStyle.Fill, Font = font, DetectUrls = false };

            btnSelectBase.Click += (_, __) => baseFile = SelectFile(txtBasePath);
            btnSelectCompare.Click += (_, __) => compareFile = SelectFile(txtComparePath);
            btnCompare.Click += (_, __) => CompareFiles();
            btnSave.Click += (_, __) => SaveResult();

            var layout = new WF.TableLayoutPanel
            {
                Dock = WF.DockStyle.Top,
                ColumnCount = 2,
                RowCount = 3,
                AutoSize = true,
                Padding = new WF.Padding(10)
            };
            layout.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.AutoSize));
            layout.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.AutoSize));

            layout.Controls.Add(btnSelectBase, 0, 0);
            layout.Controls.Add(txtBasePath, 1, 0);
            layout.Controls.Add(btnSelectCompare, 0, 1);
            layout.Controls.Add(txtComparePath, 1, 1);
            layout.Controls.Add(btnCompare, 0, 2);
            layout.Controls.Add(btnSave, 1, 2);

            Controls.Add(txtResult);
            Controls.Add(layout);
        }

        private static string SelectFile(WF.TextBox target)
        {
            using var dlg = new WF.OpenFileDialog
            {
                Filter = "Word Document (*.docx)|*.docx",
                Title = "파일 선택"
            };
            if (dlg.ShowDialog() == WF.DialogResult.OK)
            {
                target.Text = dlg.FileName;
                return dlg.FileName;
            }
            return string.Empty;
        }

        private void CompareFiles()
        {
            if (!File.Exists(baseFile) || !File.Exists(compareFile))
            {
                WF.MessageBox.Show("두 파일을 모두 선택하세요.");
                return;
            }

            var baseLines = ExtractParagraphs(baseFile);
            var compareLines = ExtractParagraphs(compareFile);
            var diffs = DiffLines(baseLines, compareLines);

            txtResult.Clear();
            foreach (var line in diffs)
            {
                if (line.StartsWith("-"))
                    txtResult.SelectionColor = SD.Color.Red;
                else if (line.StartsWith("+"))
                    txtResult.SelectionColor = SD.Color.Blue;
                else
                    txtResult.SelectionColor = SD.Color.Black;

                txtResult.AppendText(line + Environment.NewLine);
            }
        }

        private void SaveResult()
        {
            if (txtResult.TextLength == 0)
            {
                WF.MessageBox.Show("저장할 결과가 없습니다.");
                return;
            }
            using var dlg = new WF.SaveFileDialog
            {
                Filter = "Text File (*.txt)|*.txt",
                Title = "결과 저장"
            };
            if (dlg.ShowDialog() == WF.DialogResult.OK)
                File.WriteAllText(dlg.FileName, txtResult.Text);
        }

        private static List<string> ExtractParagraphs(string path)
        {
            using var doc = WordprocessingDocument.Open(path, false);
            var body = doc.MainDocumentPart?.Document.Body;
            return body == null
                ? new List<string>()
                : body.Elements<Paragraph>().Select(p => p.InnerText.Trim()).ToList();
        }

        private static IEnumerable<string> DiffLines(IList<string> a, IList<string> b)
        {
            int max = Math.Max(a.Count, b.Count);
            for (int i = 0; i < max; i++)
            {
                string lineA = i < a.Count ? a[i] : string.Empty;
                string lineB = i < b.Count ? b[i] : string.Empty;
                if (lineA == lineB)
                {
                    yield return "  " + lineA;
                }
                else
                {
                    if (!string.IsNullOrEmpty(lineA))
                        yield return "- " + lineA;
                    if (!string.IsNullOrEmpty(lineB))
                        yield return "+ " + lineB;
                }
            }
        }
    }
}
