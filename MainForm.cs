// MainForm.cs
using System;
using System.IO;               // ★ 파이썬 스크립트 경로 확인에 사용
using System.Diagnostics;      // ★ 외부 프로세스 실행
using SD = System.Drawing;
using WF = System.Windows.Forms;

namespace MyApp
{
    public class MainForm : WF.Form
    {
        // 색상
        private static readonly SD.Color COLOR_BG = SD.Color.White;
        private static readonly SD.Color COLOR_SIDEBAR = SD.Color.FromArgb(0x15, 0x65, 0xC0);
        private static readonly SD.Color COLOR_HOVER = SD.Color.FromArgb(0x19, 0x76, 0xD2);
        private static readonly SD.Color COLOR_TEXT = SD.Color.White;

        // UI
        private WF.Panel sidebar = null!;
        private WF.Panel menuContainer = null!;
        private WF.Panel mainContent = null!;   // ★ 메인 표시영역

        // 팀 목록
        private readonly string[] teams = { "영업팀", "구매팀", "설계팀", "품질팀", "생산관리팀", "시운전팀" };

        public MainForm()
        {
            Text = "Industrial Power System Team";
            StartPosition = WF.FormStartPosition.CenterScreen;
            ClientSize = new SD.Size(1000, 650);
            BackColor = COLOR_BG;

            BuildLayout();
            BuildSidebarMenus();

            // 빈 화면 안내
            ShowInMain(new WF.Label
            {
                Text = "좌측 메뉴에서 기능을 선택하세요",
                AutoSize = true,
                Font = new SD.Font("맑은 고딕", 14, SD.FontStyle.Regular),
                ForeColor = SD.Color.DimGray,
                Dock = WF.DockStyle.None
            }, center: true);
        }

        private void BuildLayout()
        {
            // 좌측 사이드바
            sidebar = new WF.Panel { Dock = WF.DockStyle.Left, Width = 220, BackColor = COLOR_SIDEBAR };
            Controls.Add(sidebar);

            // 스크롤 컨테이너
            menuContainer = new WF.Panel { Dock = WF.DockStyle.Fill, BackColor = COLOR_SIDEBAR, AutoScroll = true };
            sidebar.Controls.Add(menuContainer);

            // 우측 메인 컨텐츠
            mainContent = new WF.Panel { Dock = WF.DockStyle.Fill, BackColor = SD.Color.White };
            Controls.Add(mainContent);
            mainContent.BringToFront();
        }

        private void BuildSidebarMenus()
        {
            foreach (var team in teams) AddTeamMenu(team);
        }

        private void AddTeamMenu(string teamName)
        {
            // 상위 프레임
            var frame = new WF.Panel { Dock = WF.DockStyle.Top, Height = 40, BackColor = COLOR_SIDEBAR };
            menuContainer.Controls.Add(frame);
            frame.BringToFront();

            // 팀 라벨
            var lbl = new WF.Label
            {
                Dock = WF.DockStyle.Fill,
                Text = " - " + teamName,
                ForeColor = COLOR_TEXT,
                BackColor = SD.Color.Transparent,
                TextAlign = SD.ContentAlignment.MiddleLeft,
                Padding = new WF.Padding(16, 0, 0, 0),
                Font = new SD.Font("맑은 고딕", 11)
            };
            frame.Controls.Add(lbl);

            // 화살표
            var arrow = new WF.Label
            {
                Dock = WF.DockStyle.Right,
                Width = 28,
                Text = "▶",
                ForeColor = COLOR_TEXT,
                TextAlign = SD.ContentAlignment.MiddleCenter
            };
            frame.Controls.Add(arrow);

            // 하위 컨테이너
            var submenu = new WF.Panel { Dock = WF.DockStyle.Top, Visible = false, BackColor = COLOR_SIDEBAR };
            menuContainer.Controls.Add(submenu);
            submenu.BringToFront();

            // 하위 항목(버튼 1~N에 해당)
            int subCount = teamName == "설계팀" ? 4 : 3;
            for (int i = 1; i <= subCount; i++)
            {
                string caption =
                    teamName == "설계팀"
                    ? i switch
                    {
                        1 => "   · 물량 구분",     // QuantitySplitView
                        2 => "   · 물량 비교",     // QuantityCompareView
                        3 => "   · UNIT SIZE 구분",// ExcelUnitSizeView
                        4 => "   · 1차 전류 계산기", // CurrentCalculator
                        _ => $"   · 블럭 {i}"
                    }
                    : $"   · 블럭 {i}";

                var sub = new WF.Label
                {
                    Dock = WF.DockStyle.Top,
                    Height = 32,
                    Text = caption,
                    ForeColor = COLOR_TEXT,
                    TextAlign = SD.ContentAlignment.MiddleLeft,
                    Padding = new WF.Padding(24, 0, 0, 0),
                    Font = new SD.Font("맑은 고딕", 10),
                    BackColor = COLOR_SIDEBAR,
                    Cursor = WF.Cursors.Hand
                };

                // Hover
                sub.MouseEnter += (_, __) => sub.BackColor = COLOR_HOVER;
                sub.MouseLeave += (_, __) => sub.BackColor = COLOR_SIDEBAR;

                // Click
                int idx = i;
                sub.Click += (_, __) =>
                {
                    if (teamName == "설계팀")
                    {
                        if (idx == 1)
                        {
                            ShowInMain(new QuantitySplitView());
                        }
                        else if (idx == 2)
                        {
                            ShowInMain(new QuantityCompareView());
                        }
                        else if (idx == 3)
                        {
                            ShowInMain(new ExcelUnitSizeView());
                        }
                        else if (idx == 4)
                        {
                            TryRunPythonScript("current_calculator.py");
                        }
                    }
                    else if (teamName == "영업팀" && idx == 1)
                    {
                        // 영업팀 - 버튼1: 파이썬 스크립트 실행
                        TryRunPythonScript("myscript.py");
                    }
                    else
                    {
                        // 다른 팀/버튼은 임시 메시지
                        WF.MessageBox.Show($"{teamName} - 블럭 {idx}", "선택");
                    }
                };

                submenu.Controls.Add(sub);
                sub.BringToFront();
            }

            // Hover(상단 프레임)
            void On(object? s, EventArgs e) => frame.BackColor = COLOR_HOVER;
            void Off(object? s, EventArgs e) => frame.BackColor = COLOR_SIDEBAR;
            frame.MouseEnter += On; frame.MouseLeave += Off;
            lbl.MouseEnter += On; lbl.MouseLeave += Off;
            arrow.MouseEnter += On; arrow.MouseLeave += Off;

            // 토글
            void Toggle(object? s, EventArgs e)
            {
                submenu.Visible = !submenu.Visible;
                arrow.Text = submenu.Visible ? "▼" : "▶";
            }
            frame.Click += Toggle; lbl.Click += Toggle; arrow.Click += Toggle;
        }

        private void TryRunPythonScript(string scriptFileName = "myscript.py")
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                // 후보 경로(출력 폴더 기준 및 소스 폴더 기준 모두 탐색)
                string? scriptPath = null;
                string[] candidates =
                {
                    Path.Combine(baseDir, scriptFileName),
                    Path.Combine(baseDir, "PythonScripts", scriptFileName),
                    Path.Combine(baseDir, "pythonscripts", scriptFileName),
                };

                foreach (var c in candidates)
                {
                    if (File.Exists(c)) { scriptPath = c; break; }
                }

                // 개발 환경에서 프로젝트 루트로 역추정
                if (scriptPath == null)
                {
                    string projRoot = Path.GetFullPath(Path.Combine(baseDir, "..", "..", ".."));
                    string[] devCandidates =
                    {
                        Path.Combine(projRoot, "PythonScripts", scriptFileName),
                        Path.Combine(projRoot, "pythonscripts", scriptFileName),
                        Path.Combine(projRoot, "src", "PythonScripts", scriptFileName),
                        Path.Combine(projRoot, "src", "pythonscripts", scriptFileName),
                    };
                    foreach (var c in devCandidates)
                    {
                        if (File.Exists(c)) { scriptPath = c; break; }
                    }
                }

                if (scriptPath == null)
                {
                    WF.MessageBox.Show(
                        $"{scriptFileName}를 찾을 수 없습니다.\n" +
                        "· 파일 속성: Build Action=Content, Copy to Output Directory=Copy if newer/Always\n" +
                        $"· 경로: (출력 폴더)\\PythonScripts\\{scriptFileName} 에 존재하는지 확인하세요.",
                        "오류");
                    return;
                }

                var psi = new ProcessStartInfo
                {
                    FileName = "python",                                  // PATH 미설정 시 절대 경로로 지정
                    Arguments = $"\"{scriptPath}\"",
                    WorkingDirectory = Path.GetDirectoryName(scriptPath) ?? baseDir,
                    UseShellExecute = false,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    CreateNoWindow = true
                };

                Process.Start(psi);
            }
            catch (Exception ex)
            {
                WF.MessageBox.Show("파이썬 실행 오류: " + ex.Message, "오류");
            }
        }

        /// <summary>메인 표시 영역을 깨끗이 비우고 컨트롤을 추가합니다.</summary>
        private void ShowInMain(WF.Control control, bool center = false)
        {
            mainContent.SuspendLayout();
            mainContent.Controls.Clear();

            if (center)
            {
                // 중앙 배치용 간단한 패널
                var host = new WF.TableLayoutPanel
                {
                    Dock = WF.DockStyle.Fill,
                    ColumnCount = 3,
                    RowCount = 3
                };
                host.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.Percent, 50));
                host.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.AutoSize));
                host.ColumnStyles.Add(new WF.ColumnStyle(WF.SizeType.Percent, 50));
                host.RowStyles.Add(new WF.RowStyle(WF.SizeType.Percent, 50));
                host.RowStyles.Add(new WF.RowStyle(WF.SizeType.AutoSize));
                host.RowStyles.Add(new WF.RowStyle(WF.SizeType.Percent, 50));
                host.Controls.Add(control, 1, 1);
                mainContent.Controls.Add(host);
            }
            else
            {
                control.Dock = WF.DockStyle.Fill;
                mainContent.Controls.Add(control);
            }

            mainContent.ResumeLayout();
        }
    }
}
