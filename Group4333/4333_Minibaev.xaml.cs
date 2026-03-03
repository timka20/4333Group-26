using System;
using System.Windows;
using Group4333.Models;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;
namespace Group4333
{

    public class UserImport
    {
        [JsonPropertyName("fio")]
        public string Fio { get; set; }

        [JsonPropertyName("code_client")]
        public string CodeClient { get; set; }

        [JsonPropertyName("date_birsday")]
        public DateTime? DateBirsday { get; set; }

        [JsonPropertyName("index")]
        public string Index { get; set; }

        [JsonPropertyName("sity")]
        public string Sity { get; set; }

        [JsonPropertyName("street")]
        public string Street { get; set; }

        [JsonPropertyName("home")]
        public string Home { get; set; }

        [JsonPropertyName("kvartira")]
        public string Kvartira { get; set; }

        [JsonPropertyName("email")]
        public string Email { get; set; }
    }

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

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "JSON файл (*.json)|*.json",
                Title = "Выберите JSON файл"
            };

            if (ofd.ShowDialog() != true) return;

            try
            {
                string jsonString = File.ReadAllText(ofd.FileName);

                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    ReadCommentHandling = JsonCommentHandling.Skip,
                    AllowTrailingCommas = true
                };

                var importUsers = JsonSerializer.Deserialize<List<UserImport>>(jsonString, options);

                if (importUsers == null || importUsers.Count == 0)
                {
                    MessageBox.Show("Не удалось прочитать данные из JSON.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var optionsBuilder = new DbContextOptionsBuilder<lab3Context>();
                string connectionString = "Server=.;Database=lab3;Trusted_Connection=True;TrustServerCertificate=True;";
                optionsBuilder.UseSqlServer(connectionString);

                using (var db = new lab3Context(optionsBuilder.Options))
                {
                    var existingCodes = db.Users.Select(u => u.CodeClient).Where(c => c != null).ToHashSet();

                    int added = 0, skipped = 0, errors = 0;

                    foreach (var u in importUsers)
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(u.CodeClient))
                            {
                                skipped++;
                                continue;
                            }

                            if (existingCodes.Contains(u.CodeClient))
                            {
                                skipped++;
                                continue;
                            }

                            var user = new User
                            {
                                Fio = u.Fio,
                                CodeClient = u.CodeClient,
                                DateBirsday = u.DateBirsday,
                                Index = u.Index,
                                Sity = u.Sity,
                                Street = u.Street,
                                Home = u.Home,
                                Kvartiva = u.Kvartira,
                                Email = u.Email
                            };

                            db.Users.Add(user);
                            existingCodes.Add(u.CodeClient);
                            added++;
                        }
                        catch (Exception ex)
                        {
                            errors++;
                        }
                    }

                    db.SaveChanges();

                    MessageBox.Show($"Готово!\nДобавлено: {added}\nПропущено: {skipped}\nОшибок: {errors}",
                                   "Импорт", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var optionsBuilder = new DbContextOptionsBuilder<lab3Context>();
            string connectionString = "Server=.;Database=lab3;Trusted_Connection=True;TrustServerCertificate=True;";
            optionsBuilder.UseSqlServer(connectionString);

            dynamic wordApp = null;
            dynamic doc = null;

            try
            {
                using (var db = new lab3Context(optionsBuilder.Options))
                {
                    var users = db.Users.ToList();
                    if (!users.Any())
                    {
                        MessageBox.Show("Нет данных для экспорта.");
                        return;
                    }

                    var groupedByStreet = users.GroupBy(u => u.Street).OrderBy(g => g.Key).ToList();

                    SaveFileDialog sfd = new SaveFileDialog
                    {
                        Filter = "Документ Word (*.docx)|*.docx",
                        Title = "Сохранить отчет",
                        FileName = $"Отчет_по_улицам_{DateTime.Now:yyyyMMdd}"
                    };

                    if (sfd.ShowDialog() != true)
                        return;

                    Type wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null)
                    {
                        MessageBox.Show("Microsoft Word не установлен на этом компьютере.");
                        return;
                    }

                    wordApp = Activator.CreateInstance(wordType);
                    wordApp.Visible = false;
                    doc = wordApp.Documents.Add();

                    dynamic mainTable = doc.Tables.Add(doc.Range(0, 0), groupedByStreet.Count + 1, 2);
                    mainTable.Borders.Enable = 1;

                    mainTable.Cell(1, 1).Range.Text = "Критерий разделения на категории";
                    mainTable.Cell(1, 2).Range.Text = "Формат экспортируемых данных";
                    mainTable.Cell(1, 2).Range.Paragraphs.Alignment = 1;
                    mainTable.Cell(1, 2).Range.Font.Bold = 1;

                    int rowIndex = 2;
                    foreach (var streetGroup in groupedByStreet)
                    {
                        string streetName = streetGroup.Key ?? "Без улицы";
                        mainTable.Cell(rowIndex, 1).Range.Text = $"По улице проживания:\n{streetName}";

                        var sortedUsers = streetGroup.OrderBy(u => u.Fio).ToList();

                        dynamic subTable = doc.Tables.Add(mainTable.Cell(rowIndex, 2).Range, sortedUsers.Count + 1, 3);
                        subTable.Borders.Enable = 1;

     
                        subTable.Cell(1, 1).Range.Text = "Код клиента";
                        subTable.Cell(1, 2).Range.Text = "ФИО";
                        subTable.Cell(1, 3).Range.Text = "E-mail";

                        for (int col = 1; col <= 3; col++)
                        {
                            subTable.Cell(1, col).Range.Font.Bold = 1;
                            subTable.Cell(1, col).Range.Paragraphs.Alignment = 1;
                        }

                        int dataRow = 2;
                        foreach (var user in sortedUsers)
                        {
                            subTable.Cell(dataRow, 1).Range.Text = user.CodeClient ?? "";
                            subTable.Cell(dataRow, 2).Range.Text = user.Fio ?? "";
                            subTable.Cell(dataRow, 3).Range.Text = user.Email ?? "";
                            dataRow++;
                        }

                        subTable.AutoFitBehavior(1);
                        rowIndex++;
                    }

                    doc.SaveAs2(sfd.FileName);
                    doc.Close();
                    wordApp.Quit();

                    MessageBox.Show($"Файл сохранен:\n{sfd.FileName}", "Экспорт завершен",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка:\n{ex.Message}\n\nWord должен быть установлен.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(); } catch { }
                }
                if (wordApp != null)
                {
                    try { wordApp.Quit(); } catch { }
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}