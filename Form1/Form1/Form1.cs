using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using PdfiumViewer;

namespace SurveyDataEntry
{
    public partial class Form1 : Form
    {
        private Dictionary<string, SurveyData> surveyDict = new Dictionary<string, SurveyData>();
        private string csvFilePath = "survey_20q_backup.csv";
        private Dictionary<string, string> studentMasterDict = new Dictionary<string, string>();

        private PdfDocument pdfDocument;
        private List<int> pagesToProcess = new List<int>();

        private TableLayoutPanel mainLayout;
        private PictureBox picViewer;
        private Panel pnlInput;

        private TextBox txtStudentId = new TextBox { Width = 200, Font = new Font("맑은 고딕", 12) };
        private Label lblStudentName = new Label { Text = "대기중...", Font = new Font("맑은 고딕", 11, FontStyle.Bold), ForeColor = Color.Gray, AutoSize = true };

        private TextBox txtAnswers = new TextBox { Width = 280, Font = new Font("맑은 고딕", 14), MaxLength = 20 };

        private TextBox txtQ1Other = new TextBox { Width = 280, Font = new Font("맑은 고딕", 10), Enabled = false, BackColor = Color.LightGray };
        private TextBox txtQ11_1 = new TextBox { Width = 280, Font = new Font("맑은 고딕", 10) };
        private TextBox txtQ11_2 = new TextBox { Width = 280, Font = new Font("맑은 고딕", 10) };
        private TextBox txtQ16Other = new TextBox { Width = 280, Font = new Font("맑은 고딕", 10), Enabled = false, BackColor = Color.LightGray };
        private TextBox txtQ18_1 = new TextBox { Width = 280, Font = new Font("맑은 고딕", 10) };
        private TextBox txtQ18_2 = new TextBox { Width = 280, Font = new Font("맑은 고딕", 10) };

        private Label lblStatus = new Label { Width = 300, ForeColor = Color.Blue, Font = new Font("맑은 고딕", 11, FontStyle.Bold) };
        private Label lblImageProgress = new Label { Width = 300, ForeColor = Color.DarkGreen, Font = new Font("맑은 고딕", 14, FontStyle.Bold) };
        private Button btnExportExcel = new Button { Width = 280, Height = 40, Text = "최종 엑셀 추출", Font = new Font("맑은 고딕", 12, FontStyle.Bold) };

        // ★ 요약창 크기를 키우고 (Height 290) 폰트 정렬을 다듬었습니다.
        private GroupBox grpRecent = new GroupBox { Text = "방금 저장된 데이터 요약", Width = 280, Height = 310, Font = new Font("맑은 고딕", 10, FontStyle.Bold), ForeColor = Color.DarkSlateGray };
        private Label lblRecentInfo = new Label { Text = "학번: -\n이름: -\n답안: -\n주관식 문항: 없음", AutoSize = true, Font = new Font("맑은 고딕", 9), Location = new Point(12, 25), ForeColor = Color.Black };

        public Form1()
        {
            SetupUI();
            this.Shown += Form1_Shown;
        }

        private void Form1_Load(object sender, EventArgs e) { }
        protected override void OnFormClosed(FormClosedEventArgs e) { pdfDocument?.Dispose(); base.OnFormClosed(e); }

        private void Form1_Shown(object sender, EventArgs e)
        {
            MessageBox.Show(this, "먼저 [학생 명단 CSV 파일]을 선택해주세요.", "1단계 (명단 로드)");
            LoadStudentMasterData();

            MessageBox.Show(this, "이제 [설문지 PDF 파일]을 선택해주세요.", "2단계 (PDF 로드)");
            LoadPdfOnStartup();
        }

        private void SetupUI()
        {
            this.Text = "20문항 하이브리드 설문 (객관식 상세 매칭 요약 기능)";
            this.Size = new Size(1600, 900);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.WindowState = FormWindowState.Maximized;

            mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30f));

            picViewer = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(240, 240, 240), Cursor = Cursors.Hand };
            picViewer.MouseClick += (s, e) => { SkipCurrentPage(); };

            pnlInput = new Panel { Dock = DockStyle.Fill, Padding = new Padding(15), AutoScroll = true };
            int yPos = 10;

            pnlInput.Controls.Add(lblImageProgress); lblImageProgress.Location = new Point(10, yPos); yPos += 35;

            pnlInput.Controls.Add(new Label { Text = "0. 학번:", AutoSize = true, Font = new Font("맑은 고딕", 10, FontStyle.Bold), Location = new Point(10, yPos) }); yPos += 25;
            pnlInput.Controls.Add(txtStudentId); txtStudentId.Location = new Point(10, yPos);
            pnlInput.Controls.Add(lblStudentName); lblStudentName.Location = new Point(220, yPos + 3); yPos += 35;

            pnlInput.Controls.Add(new Label { Text = "1. 객관식 20개 연속입력 (스페이스바=스킵)", AutoSize = true, ForeColor = Color.DarkRed, Font = new Font("맑은 고딕", 10, FontStyle.Bold), Location = new Point(10, yPos) }); yPos += 25;
            pnlInput.Controls.Add(txtAnswers); txtAnswers.Location = new Point(10, yPos); yPos += 45;

            pnlInput.Controls.Add(new Label { Text = "▶ 1번 기타 사유 (5번 선택시 활성):", AutoSize = true, Font = new Font("맑은 고딕", 9), Location = new Point(10, yPos) }); yPos += 20;
            pnlInput.Controls.Add(txtQ1Other); txtQ1Other.Location = new Point(10, yPos); yPos += 35;

            pnlInput.Controls.Add(new Label { Text = "▶ 11-1번 답변:", AutoSize = true, Font = new Font("맑은 고딕", 9), Location = new Point(10, yPos) }); yPos += 20;
            pnlInput.Controls.Add(txtQ11_1); txtQ11_1.Location = new Point(10, yPos); yPos += 35;

            pnlInput.Controls.Add(new Label { Text = "▶ 11-2번 답변:", AutoSize = true, Font = new Font("맑은 고딕", 9), Location = new Point(10, yPos) }); yPos += 20;
            pnlInput.Controls.Add(txtQ11_2); txtQ11_2.Location = new Point(10, yPos); yPos += 45;

            pnlInput.Controls.Add(new Label { Text = "▶ 16번 기타 사유 (7번 선택시 활성):", AutoSize = true, Font = new Font("맑은 고딕", 9), Location = new Point(10, yPos) }); yPos += 20;
            pnlInput.Controls.Add(txtQ16Other); txtQ16Other.Location = new Point(10, yPos); yPos += 35;

            pnlInput.Controls.Add(new Label { Text = "▶ 18-1번 답변:", AutoSize = true, Font = new Font("맑은 고딕", 9), Location = new Point(10, yPos) }); yPos += 20;
            pnlInput.Controls.Add(txtQ18_1); txtQ18_1.Location = new Point(10, yPos); yPos += 35;

            pnlInput.Controls.Add(new Label { Text = "▶ 18-2번 답변:", AutoSize = true, Font = new Font("맑은 고딕", 9), Location = new Point(10, yPos) }); yPos += 20;
            pnlInput.Controls.Add(txtQ18_2); txtQ18_2.Location = new Point(10, yPos); yPos += 45;

            pnlInput.Controls.Add(lblStatus); lblStatus.Location = new Point(10, yPos); yPos += 30;
            pnlInput.Controls.Add(btnExportExcel); btnExportExcel.Location = new Point(10, yPos); yPos += 55;

            grpRecent.Controls.Add(lblRecentInfo);
            pnlInput.Controls.Add(grpRecent); grpRecent.Location = new Point(10, yPos);

            txtStudentId.KeyDown += FocusNext_KeyDown;
            txtAnswers.KeyDown += FocusNext_KeyDown;
            txtQ1Other.KeyDown += FocusNext_KeyDown;
            txtQ11_1.KeyDown += FocusNext_KeyDown;
            txtQ11_2.KeyDown += FocusNext_KeyDown;
            txtQ16Other.KeyDown += FocusNext_KeyDown;
            txtQ18_1.KeyDown += FocusNext_KeyDown;
            txtQ18_2.KeyDown += FocusNext_KeyDown;

            txtAnswers.TextChanged += TxtAnswers_TextChanged;
            txtAnswers.KeyPress += OnlyNumbers_KeyPress;

            btnExportExcel.Click += BtnExportExcel_Click;

            mainLayout.Controls.Add(picViewer, 0, 0);
            mainLayout.Controls.Add(pnlInput, 1, 0);
            this.Controls.Add(mainLayout);
        }

        private void TxtAnswers_TextChanged(object sender, EventArgs e)
        {
            string ans = txtAnswers.Text;
            if (ans.Length >= 1 && ans[0] == '5') { txtQ1Other.Enabled = true; txtQ1Other.BackColor = Color.White; }
            else { txtQ1Other.Enabled = false; txtQ1Other.BackColor = Color.LightGray; txtQ1Other.Clear(); }

            if (ans.Length >= 16 && ans[15] == '7') { txtQ16Other.Enabled = true; txtQ16Other.BackColor = Color.White; }
            else { txtQ16Other.Enabled = false; txtQ16Other.BackColor = Color.LightGray; txtQ16Other.Clear(); }
        }

        private void OnlyNumbers_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == ' ') e.KeyChar = '-';
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != 8 && e.KeyChar != 13 && e.KeyChar != '-') e.Handled = true;
        }

        private void FocusNext_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                Control current = sender as Control;

                if (current == txtStudentId)
                {
                    string id = txtStudentId.Text.Trim();
                    if (studentMasterDict.Count > 0)
                    {
                        if (studentMasterDict.ContainsKey(id)) { lblStudentName.Text = studentMasterDict[id]; lblStudentName.ForeColor = Color.Blue; txtAnswers.Focus(); }
                        else { lblStudentName.Text = "명단 없음!"; lblStudentName.ForeColor = Color.Red; txtStudentId.SelectAll(); }
                    }
                    else txtAnswers.Focus();
                }
                else if (current == txtAnswers)
                {
                    if (txtAnswers.Text.Length < 20) { MessageBox.Show("객관식 답안 20자리를 모두 채워주세요."); return; }
                    if (txtQ1Other.Enabled) txtQ1Other.Focus();
                    else txtQ11_1.Focus();
                }
                else if (current == txtQ1Other) { txtQ11_1.Focus(); }
                else if (current == txtQ11_1) { txtQ11_2.Focus(); }
                else if (current == txtQ11_2)
                {
                    if (txtQ16Other.Enabled) txtQ16Other.Focus();
                    else txtQ18_1.Focus();
                }
                else if (current == txtQ16Other) { txtQ18_1.Focus(); }
                else if (current == txtQ18_1) { txtQ18_2.Focus(); }
                else if (current == txtQ18_2) { SaveCurrentData(); }
            }
        }

        private void SaveCurrentData()
        {
            string id = txtStudentId.Text.Trim();
            if (string.IsNullOrEmpty(id)) return;

            var data = new SurveyData
            {
                StudentId = id,
                Answers = txtAnswers.Text,
                Q1_Other = txtQ1Other.Text.Trim(),
                Q11_1 = txtQ11_1.Text.Trim(),
                Q11_2 = txtQ11_2.Text.Trim(),
                Q16_Other = txtQ16Other.Text.Trim(),
                Q18_1 = txtQ18_1.Text.Trim(),
                Q18_2 = txtQ18_2.Text.Trim()
            };

            surveyDict[id] = data;
            string name = studentMasterDict.ContainsKey(id) ? studentMasterDict[id] : "알수없음";

            using (StreamWriter sw = new StreamWriter(csvFilePath, true, Encoding.UTF8))
            {
                sw.WriteLine($"{data.StudentId},\"{data.Answers}\",\"{data.Q1_Other}\",\"{data.Q11_1}\",\"{data.Q11_2}\",\"{data.Q16_Other}\",\"{data.Q18_1}\",\"{data.Q18_2}\"");
            }

            // ★ 핵심 개편: 객관식 번호를 5개씩 슬라이스해서 시각적으로 매칭해주는 로직
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"■ 학번/성명: {id} ({name})");
            sb.AppendLine($"■ 객관식 상세내역:");

            // 20자리 숫자를 돌면서 번호와 매칭
            for (int i = 0; i < data.Answers.Length; i++)
            {
                // 두 자리 숫자로 맞춤 (예: 01번: 3, 02번: 4)
                sb.Append($"{i + 1:D2}번: {data.Answers[i]}   ");

                // 5개마다 줄바꿈해서 보기 좋게 정렬
                if ((i + 1) % 5 == 0) sb.AppendLine();
            }

            sb.AppendLine($"■ 1번사유 : {(string.IsNullOrEmpty(data.Q1_Other) ? "-" : data.Q1_Other)}");
            sb.AppendLine($"■ 11번답 : {data.Q11_1} / {data.Q11_2}");
            sb.AppendLine($"■ 16번사유 : {(string.IsNullOrEmpty(data.Q16_Other) ? "-" : data.Q16_Other)}");
            sb.Append($"■ 18번답 : {data.Q18_1} / {data.Q18_2}");

            lblRecentInfo.Text = sb.ToString();
            lblStatus.Text = $"[{name}] 저장 완료!";

            txtStudentId.Clear(); txtAnswers.Clear(); txtQ1Other.Clear(); txtQ11_1.Clear(); txtQ11_2.Clear(); txtQ16Other.Clear(); txtQ18_1.Clear(); txtQ18_2.Clear();
            lblStudentName.Text = "대기중..."; lblStudentName.ForeColor = Color.Gray;
            txtStudentId.Focus();

            SkipCurrentPage();
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
        {
            if (surveyDict.Count == 0) { MessageBox.Show("데이터가 없습니다."); return; }
            OpenFileDialog ofd = new OpenFileDialog { Title = "원본 엑셀 템플릿 선택", Filter = "Excel Files|*.xlsx" };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string outputPath = Path.Combine(Path.GetDirectoryName(ofd.FileName), "20문항_수요조사결과_완료.xlsx");
                try
                {
                    using (var workbook = new XLWorkbook(ofd.FileName))
                    {
                        var ws = workbook.Worksheet(1);
                        int lastRow = ws.LastRowUsed().RowNumber();

                        for (int row = 9; row <= lastRow; row++)
                        {
                            string id = ws.Cell(row, 3).GetString().Trim();
                            if (surveyDict.ContainsKey(id))
                            {
                                var data = surveyDict[id];
                                ws.Cell(row, 9).Value = "o";

                                for (int i = 0; i < data.Answers.Length; i++)
                                {
                                    char ansChar = data.Answers[i];
                                    if (ansChar != '-') ws.Cell(row, 11 + i).Value = ansChar.ToString();
                                }

                                int otherStartCol = 11 + 20;
                                ws.Cell(row, otherStartCol).Value = data.Q1_Other;
                                ws.Cell(row, otherStartCol + 1).Value = data.Q11_1;
                                ws.Cell(row, column: otherStartCol + 2).Value = data.Q11_2;
                                ws.Cell(row, column: otherStartCol + 3).Value = data.Q16_Other;
                                ws.Cell(row, column: otherStartCol + 4).Value = data.Q18_1;
                                ws.Cell(row, column: otherStartCol + 5).Value = data.Q18_2;
                            }
                        }
                        workbook.SaveAs(outputPath);
                    }
                    MessageBox.Show($"변환 완료!\n저장위치: {outputPath}");
                }
                catch (Exception ex) { MessageBox.Show("에러: " + ex.Message); }
            }
        }

        private void LoadStudentMasterData()
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Title = "학생 명단 CSV 파일 선택", Filter = "CSV Files|*.csv" })
            {
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    try { foreach (var line in File.ReadAllLines(ofd.FileName, Encoding.Default)) { var parts = line.Split(','); if (parts.Length > 3) { string id = parts[2].Trim(); if (id.Length >= 8 && id.All(char.IsDigit)) studentMasterDict[id] = parts[3].Trim(); } } } catch { }
                }
            }
        }
        private void LoadPdfOnStartup()
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Title = "설문지 PDF 파일 선택", Filter = "PDF Files|*.pdf" })
            {
                if (ofd.ShowDialog(this) == DialogResult.OK) { pdfDocument = PdfDocument.Load(ofd.FileName); pagesToProcess = Enumerable.Range(0, pdfDocument.PageCount).ToList(); ShowCurrentPage(); }
            }
        }
        private void ShowCurrentPage()
        {
            if (picViewer.Image != null) { picViewer.Image.Dispose(); picViewer.Image = null; }
            if (pagesToProcess.Count == 0) { lblImageProgress.Text = "모든 작업 완료!"; return; }
            try { picViewer.Image = pdfDocument.Render(pagesToProcess[0], (int)(pdfDocument.PageSizes[pagesToProcess[0]].Width * 150 / 72.0), (int)(pdfDocument.PageSizes[pagesToProcess[0]].Height * 150 / 72.0), 150, 150, false); lblImageProgress.Text = $"현재 페이지: {pagesToProcess[0] + 1} / {pdfDocument.PageCount}"; } catch { }
        }
        private void SkipCurrentPage() { try { if (pagesToProcess.Count > 0) { pagesToProcess.RemoveAt(0); ShowCurrentPage(); } } catch { } }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) { if (keyData == Keys.Escape) { SkipCurrentPage(); return true; } return base.ProcessCmdKey(ref msg, keyData); }
    }

    public class SurveyData
    {
        public string StudentId { get; set; }
        public string Answers { get; set; }
        public string Q1_Other { get; set; }
        public string Q11_1 { get; set; }
        public string Q11_2 { get; set; }
        public string Q16_Other { get; set; }
        public string Q18_1 { get; set; }
        public string Q18_2 { get; set; }
    }
}