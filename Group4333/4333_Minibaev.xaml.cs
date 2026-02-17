using System;
using System.Windows;
using Group4333.Models;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic; 

namespace Group4333
{
    public partial class _4333_Minibaev : Window
    {
        public _4333_Minibaev()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (ofd.ShowDialog() != true)
                return;

            Excel.Application ObjWorkExcel = null;
            Excel.Workbook ObjWorkBook = null;
            Excel.Worksheet ObjWorkSheet = null;

            ObjWorkExcel = new Excel.Application();
            ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            string[,] list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            var optionsBuilder = new Microsoft.EntityFrameworkCore.DbContextOptionsBuilder<lab3Context>();
            string connectionString = "Server=.;Database=lab3;Trusted_Connection=True;TrustServerCertificate=True;";
            optionsBuilder.UseSqlServer(connectionString);

            

            using (var db = new lab3Context(optionsBuilder.Options))
            {
                var addedCodes = new HashSet<string>();
                int skippedCount = 0;

                for (int i = 1; i < _rows; i++)
                {
                    string codeClient = list[i, 1];

                    if (string.IsNullOrWhiteSpace(codeClient))
                        continue;

                    if (addedCodes.Contains(codeClient))
                    {
                        skippedCount++;
                        continue; 
                    }

                    var user = new User()
                    {
                        Fio = list[i, 0],
                        CodeClient = codeClient,
                        DateBirsday = DateTime.TryParse(list[i, 2], out var dt) ? dt : null,
                        Index = list[i, 3],
                        Sity = list[i, 4],
                        Street = list[i, 5],
                        Home = list[i, 6],
                        Kvartiva = list[i, 7],
                        Email = list[i, 8]
                    };

                    db.Users.Add(user);
                    addedCodes.Add(codeClient); 
                }

                db.SaveChanges();

                string msg = "Данные успешно добавлены.";
                if (skippedCount > 0)
                    msg += $"\n(Пропущено дубликатов: {skippedCount})";

                MessageBox.Show(msg);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var optionsBuilder = new Microsoft.EntityFrameworkCore.DbContextOptionsBuilder<lab3Context>();
            string connectionString = "Server=.;Database=lab3;Trusted_Connection=True;TrustServerCertificate=True;";
            optionsBuilder.UseSqlServer(connectionString);

            using (var db = new lab3Context(optionsBuilder.Options))
            {
                var allUsers = db.Users.ToList();
                var usersByStreet = allUsers.GroupBy(u => u.Street).OrderBy(g => g.Key).ToList();

                if (!usersByStreet.Any())
                {
                    MessageBox.Show("В базе данных нет пользователей для экспорта.");
                    return;
                }

                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false; 
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                while (workbook.Sheets.Count > 1)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];
                    sheet.Delete();
                }

                int sheetIndex = 0;

                foreach (var streetGroup in usersByStreet)
                {
                    string streetName = streetGroup.Key ?? "Без улицы";
                    if (sheetIndex > 0)
                    {
                        Excel.Worksheet newSheet = (Excel.Worksheet)workbook.Sheets.Add(
                            After: workbook.Sheets[workbook.Sheets.Count]);
                        newSheet.Name = streetName.Length > 31 ? streetName.Substring(0, 31) : streetName;
                    }
                    else
                    {
                        Excel.Worksheet firstSheet = (Excel.Worksheet)workbook.Sheets[1];
                        firstSheet.Name = streetName.Length > 31 ? streetName.Substring(0, 31) : streetName;
                    }

                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[sheetIndex + 1];

                    worksheet.Cells[1, 1] = "Код клиента";
                    worksheet.Cells[1, 2] = "ФИО";
                    worksheet.Cells[1, 3] = "E-mail";
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 3]];
                    headerRange.Font.Bold = true;

                    var sortedUsers = streetGroup.OrderBy(u => u.Fio).ToList();

                    int startRow = 2;
                    foreach (var user in sortedUsers)
                    {
                        worksheet.Cells[startRow, 1] = user.CodeClient;
                        worksheet.Cells[startRow, 2] = user.Fio;
                        worksheet.Cells[startRow, 3] = user.Email ?? "";
                        startRow++;
                    }

                    if (startRow > 2)
                    {
                        Excel.Range dataRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[startRow - 1, 3]];
                        dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        dataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    }

                    worksheet.Columns.AutoFit();

                    sheetIndex++;
                }

                excelApp.Visible = true;
                excelApp.UserControl = true;
            }
        }
    }
}