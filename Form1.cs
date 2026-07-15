using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OfficeOpenXml;

namespace PRAK1
{
    public class Form1 : Form
    {
        private DataGridView dgvStudents;
        private Button btnLoadExcel;
        private Button btnCalculate;
        private Button btnCreateExample;
        private OpenFileDialog openFileDialog;
        private List<Student> students = new List<Student>();
        private readonly string[] subjects = { "Математика", "Физика", "Программирование", "История", "Физкультура", "Английский" };

        public Form1()
        {
            InitializeComponent();
            CheckAndLoadDefaultFile();
        }

        private void InitializeComponent()
        {
            this.Text = "Расчёт стипендий студентам";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.WhiteSmoke;

            dgvStudents = new DataGridView
            {
                Location = new Point(20, 20),
                Size = new Size(840, 400),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            btnLoadExcel = new Button
            {
                Text = "📂 Открыть Excel",
                Location = new Point(20, 440),
                Size = new Size(150, 40),
                BackColor = Color.LightBlue,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnLoadExcel.Click += BtnLoadExcel_Click;

            btnCalculate = new Button
            {
                Text = "💰 Рассчитать стипендии",
                Location = new Point(190, 440),
                Size = new Size(180, 40),
                BackColor = Color.LightGreen,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnCalculate.Click += BtnCalculate_Click;

            btnCreateExample = new Button
            {
                Text = "📄 Создать пример Excel",
                Location = new Point(390, 440),
                Size = new Size(180, 40),
                BackColor = Color.LightYellow,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            btnCreateExample.Click += BtnCreateExample_Click;

            openFileDialog = new OpenFileDialog
            {
                Filter = "Excel файлы|*.xlsx",
                Title = "Выберите файл с данными студентов"
            };

            Controls.Add(dgvStudents);
            Controls.Add(btnLoadExcel);
            Controls.Add(btnCalculate);
            Controls.Add(btnCreateExample);
        }

        private void CheckAndLoadDefaultFile()
        {
            string defaultPath = Path.Combine(Application.StartupPath, "Students.xlsx");
            if (!File.Exists(defaultPath))
                CreateExampleExcelFile(defaultPath);
            LoadStudentsFromExcel(defaultPath);
        }

        private void CreateExampleExcelFile(string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Студенты");
                worksheet.Cells[1, 1].Value = "ФИО";
                worksheet.Cells[1, 2].Value = "Многодетная (да/нет)";
                for (int i = 0; i < subjects.Length; i++)
                    worksheet.Cells[1, i + 3].Value = subjects[i];

                var sampleStudents = new (string name, bool largeFamily, int[] grades)[]
                {
                    ("Иванов Иван", true, new int[] {5,4,5,5,4,5}),
                    ("Петрова Мария", false, new int[] {4,4,3,4,5,4}),
                    ("Сидоров Алексей", true, new int[] {3,3,4,3,4,3}),
                    ("Козлова Анна", false, new int[] {5,5,5,5,5,5}),
                    ("Николаев Дмитрий", false, new int[] {2,3,3,4,3,2})
                };

                for (int row = 0; row < sampleStudents.Length; row++)
                {
                    var s = sampleStudents[row];
                    worksheet.Cells[row + 2, 1].Value = s.name;
                    worksheet.Cells[row + 2, 2].Value = s.largeFamily ? "да" : "нет";
                    for (int col = 0; col < s.grades.Length; col++)
                        worksheet.Cells[row + 2, col + 3].Value = s.grades[col];
                }

                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                package.SaveAs(new FileInfo(filePath));
            }
        }

        private void LoadStudentsFromExcel(string filePath)
        {
            students.Clear();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int row = 2;
                while (worksheet.Cells[row, 1].Value != null)
                {
                    string name = worksheet.Cells[row, 1].Value.ToString();
                    string largeFamilyStr = worksheet.Cells[row, 2].Value?.ToString().ToLower() ?? "нет";
                    bool isLargeFamily = (largeFamilyStr == "да" || largeFamilyStr == "yes");

                    List<int> grades = new List<int>();
                    for (int col = 0; col < subjects.Length; col++)
                    {
                        object val = worksheet.Cells[row, col + 3].Value;
                        if (val != null && int.TryParse(val.ToString(), out int grade))
                            grades.Add(grade);
                        else
                            grades.Add(0);
                    }
                    students.Add(new Student { Name = name, IsLargeFamily = isLargeFamily, Grades = grades });
                    row++;
                }
            }
            RefreshDataGridView();
        }

        private void RefreshDataGridView()
        {
            dgvStudents.Columns.Clear();
            dgvStudents.Columns.Add("Name", "ФИО");
            dgvStudents.Columns.Add("LargeFamily", "Многодетная");
            foreach (var subj in subjects)
                dgvStudents.Columns.Add(subj, subj);
            dgvStudents.Columns.Add("AvgGrade", "Ср. балл");
            dgvStudents.Columns.Add("Status", "Тип стипендии");
            dgvStudents.Columns.Add("Amount", "Сумма (руб.)");

            dgvStudents.Rows.Clear();
            foreach (var s in students)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgvStudents);
                row.Cells[0].Value = s.Name;
                row.Cells[1].Value = s.IsLargeFamily ? "Да" : "Нет";
                for (int i = 0; i < subjects.Length; i++)
                    row.Cells[i + 2].Value = s.Grades[i];
                dgvStudents.Rows.Add(row);
            }
        }

        private void BtnCalculate_Click(object sender, EventArgs e)
        {
            if (students.Count == 0) return;

            foreach (DataGridViewRow row in dgvStudents.Rows)
            {
                if (row.IsNewRow) continue;
                string studentName = row.Cells["Name"].Value.ToString();
                Student student = students.FirstOrDefault(s => s.Name == studentName);
                if (student == null) continue;

                IScholarshipCalculator calculator;
                if (student.IsLargeFamily)
                    calculator = new SocialScholarship(true);
                else
                    calculator = new StandardScholarship();

                double avg = calculator.CalculateAverageGrade(student.Grades);
                int amount = calculator.DetermineScholarship(avg);
                string status = calculator.GetStudentStatus(avg);

                row.Cells["AvgGrade"].Value = avg.ToString("F2");
                row.Cells["Status"].Value = status;
                row.Cells["Amount"].Value = amount;
            }
        }

        private void BtnLoadExcel_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                LoadStudentsFromExcel(openFileDialog.FileName);
        }

        private void BtnCreateExample_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel файлы|*.xlsx",
                FileName = "Students.xlsx"
            };
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                CreateExampleExcelFile(saveDialog.FileName);
                MessageBox.Show("Файл примера создан!");
                LoadStudentsFromExcel(saveDialog.FileName);
            }
        }
    }
}