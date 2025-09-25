using System;
using System.Collections.Generic;
using System.Windows.Forms;
using SD = System.Drawing;

namespace MyApp
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

            var (passText, mismatchText, elbText) = ExcelUnitSizeAnalysisService.Analyze(ofd.FileName);

            _txtPass.Text = string.IsNullOrWhiteSpace(passText) ? "통과 항목 없음" : passText;
            _txtMismatch.Text = string.IsNullOrWhiteSpace(mismatchText) ? "불일치 없음" : mismatchText;
            _txtElb.Text = string.IsNullOrWhiteSpace(elbText) ? "ELB 없음" : elbText;

            ResetSearchState(clearKeyword: false);
            _txtSearch.Focus();
        }
    }
}
