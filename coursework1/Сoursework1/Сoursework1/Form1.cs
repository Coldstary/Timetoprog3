

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms.DataVisualization.Charting;

namespace Сoursework1
{
    public partial class Tablecreator : Form
    {

        private Panel chartPanel;
        private Button btnCloseChart;

        public Tablecreator()
        {
            InitializeComponent();

        }


        private void Tablecreator_Load(object sender, EventArgs e)
        {
            ConfigureDataGridView();

        }

        private void ConfigureDataGridView()
        {
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.ReadOnly = true;
        }

        private void CreateChartBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Documents (*.docx, *.doc)|*.docx;*.doc|All files (*.*)|*.*";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Выберите Word документы";
            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;

            try
            {
                openFileDialog.FileName = string.Empty;
                DialogResult result = openFileDialog.ShowDialog(this);

                if (result == DialogResult.OK)
                {
                    if (openFileDialog.FileNames.Length > 0)
                    {
                        UpdateStatus(string.Format("Обработка {0} файлов...", openFileDialog.FileNames.Length));
                        ProcessWordFiles(openFileDialog.FileNames);
                    }
                    else
                    {
                        MessageBox.Show("Файлы не выбраны", "Информация",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    UpdateStatus("Выбор файлов отменен");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при выборе файлов: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                openFileDialog.Dispose();
            }
        }

        private void UpdateStatus(string message)
        {
            if (lblStatus != null && !lblStatus.IsDisposed)
            {
                lblStatus.Text = message;
            }
        }

        private void ProcessWordFiles(string[] filePaths)
        {
            DataTable displayTable = CreateDisplayDataTable();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            int totalTables = 0;
            bool firstFile = true;

            try
            {
                foreach (string filePath in filePaths)
                {
                    try
                    {
                        if (!firstFile)
                        {
                            AddSeparatorRow(displayTable);
                        }

                        int tablesProcessed = ProcessSingleWordFile(filePath, displayTable);
                        totalTables += tablesProcessed;
                        firstFile = false;
                    }
                    catch (Exception ex)
                    {
                        AddErrorRow(displayTable, Path.GetFileName(filePath), string.Format("Ошибка: {0}", ex.Message));
                    }
                }

                if (displayTable.Rows.Count > 0)
                {
                    dataGridView1.DataSource = displayTable;
                    UpdateStatus(string.Format("Обработано {0} файлов, найдено {1} таблиц", filePaths.Length, totalTables));
                }
                else
                {
                    UpdateStatus("Таблицы не найдены в выбранных файлах");
                    MessageBox.Show("В выбранных файлах не найдено таблиц", "Информация",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Критическая ошибка: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable CreateDisplayDataTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Файл", typeof(string));
            table.Columns.Add("Номер таблицы", typeof(string));
            table.Columns.Add("Семестр", typeof(string));
            table.Columns.Add("Колонка 1", typeof(string));
            table.Columns.Add("Колонка 2", typeof(string));
            table.Columns.Add("Колонка 3", typeof(string));
            table.Columns.Add("Колонка 4", typeof(string));
            table.Columns.Add("Колонка 5", typeof(string));
            table.Columns.Add("Колонка 6", typeof(string));
            table.Columns.Add("Колонка 7", typeof(string));
            table.Columns.Add("Колонка 8", typeof(string));
            return table;
        }

        private void AddSeparatorRow(DataTable displayTable)
        {
            DataRow separatorRow = displayTable.NewRow();
            for (int i = 0; i < displayTable.Columns.Count; i++)
            {
                separatorRow[i] = "---";
            }
            displayTable.Rows.Add(separatorRow);
        }

        private int ProcessSingleWordFile(string filePath, DataTable displayTable)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;
            int tablesCount = 0;
            string fileName = Path.GetFileName(filePath);

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                doc = wordApp.Documents.Open(
                    FileName: filePath,
                    ReadOnly: true,
                    Visible: false
                );

                // Извлекаем заголовок документа
                string headerText = ExtractDocumentHeader(doc);

                // Извлекаем подпись преподавателя
                string signatureText = ExtractTeacherSignature(doc);

                // Добавляем заголовок как первую строку в таблицу
                if (!string.IsNullOrEmpty(headerText))
                {
                    DataRow headerRow = displayTable.NewRow();
                    headerRow["Файл"] = fileName;
                    headerRow["Номер таблицы"] = "Документ";
                    headerRow["Семестр"] = "";
                    headerRow["Колонка 1"] = headerText;
                    displayTable.Rows.Add(headerRow);
                }

                // Обрабатываем ВСЕ таблицы в документе (всегда 4 таблицы)
                for (int i = 1; i <= doc.Tables.Count; i++)
                {
                    try
                    {
                        Word.Table wordTable = doc.Tables[i];
                        string semester = DetermineSemester(wordTable, i, doc.Tables.Count);

                        // Обрабатываем таблицу с правильным определением семестра
                        ProcessWordTable(wordTable, displayTable, fileName, i, semester);
                        tablesCount++;

                        Marshal.ReleaseComObject(wordTable);
                    }
                    catch (Exception tableEx)
                    {
                        AddErrorRow(displayTable, fileName, string.Format("Ошибка в таблице {0}: {1}", i, tableEx.Message));
                    }
                }

                // Добавляем подпись преподавателя
                if (!string.IsNullOrEmpty(signatureText))
                {
                    DataRow signatureRow = displayTable.NewRow();
                    signatureRow["Файл"] = fileName;
                    signatureRow["Номер таблицы"] = "Подпись";
                    signatureRow["Семестр"] = "";
                    signatureRow["Колонка 1"] = signatureText;
                    displayTable.Rows.Add(signatureRow);
                }

                if (doc.Tables.Count == 0)
                {
                    AddErrorRow(displayTable, fileName, "Таблицы не найдены");
                }
            }
            finally
            {
                if (doc != null)
                {
                    try
                    {
                        doc.Close(false);
                        Marshal.ReleaseComObject(doc);
                    }
                    catch { }
                }

                if (wordApp != null)
                {
                    try
                    {
                        wordApp.Quit(false);
                        Marshal.ReleaseComObject(wordApp);
                    }
                    catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return tablesCount;
        }

        // Метод для извлечения заголовка документа
        private string ExtractDocumentHeader(Word.Document doc)
        {
            try
            {
                string fullText = doc.Content.Text;

                // Ищем начало заголовка
                int startIndex = fullText.IndexOf("ИНДИВИДУАЛЬНЫЙ ПЛАН ПРЕПОДАВАТЕЛЯ", StringComparison.OrdinalIgnoreCase);
                if (startIndex >= 0)
                {
                    // Ищем конец заголовка (до следующего заголовка или конца текста)
                    int endIndex = startIndex;

                    // Ищем конец строки или следующую значимую часть
                    for (int i = startIndex; i < Math.Min(startIndex + 500, fullText.Length); i++)
                    {
                        if (fullText[i] == '\r' || fullText[i] == '\n')
                        {
                            // Проверяем, не начинается ли следующая строка с новой информации
                            int nextLineStart = i + 1;
                            while (nextLineStart < fullText.Length && (fullText[nextLineStart] == '\r' || fullText[nextLineStart] == '\n'))
                            {
                                nextLineStart++;
                            }

                            // Если следующая строка начинается с новой информации, заканчиваем заголовок
                            if (nextLineStart < fullText.Length)
                            {
                                string nextChar = fullText.Substring(nextLineStart, Math.Min(20, fullText.Length - nextLineStart));
                                if (nextChar.Contains("Федеральное") || nextChar.Contains("Министерство") ||
                                    nextChar.Contains("1.") || nextChar.Contains("1-й"))
                                {
                                    endIndex = i;
                                    break;
                                }
                            }
                        }
                        endIndex = i;
                    }

                    string headerText = fullText.Substring(startIndex, endIndex - startIndex);
                    return CleanText(headerText);
                }
            }
            catch (Exception)
            {
                // В случае ошибки возвращаем пустую строку
            }

            return "";
        }

        // Метод для извлечения подписи преподавателя
        private string ExtractTeacherSignature(Word.Document doc)
        {
            try
            {
                string fullText = doc.Content.Text;

                // Ищем подпись преподавателя (обычно в конце документа)
                string[] lines = fullText.Split(new[] { "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                // Ищем с конца
                for (int i = lines.Length - 1; i >= 0; i--)
                {
                    string line = lines[i].Trim();
                    if (line.StartsWith("Преподаватель", StringComparison.OrdinalIgnoreCase))
                    {
                        return CleanText(line);
                    }
                }
            }
            catch (Exception)
            {
                // В случае ошибки возвращаем пустую строку
            }
            return "";
        }

        private string DetermineSemester(Word.Table wordTable, int tableIndex, int totalTables)
        {
            // В каждом документе всегда 4 таблицы
            // Таблица 1 и 2: без семестра (заголовочные)
            // Таблица 3: 1 семестр
            // Таблица 4: 2 семестр

            if (tableIndex == 1 || tableIndex == 2)
            {
                return ""; // Пустая строка для заголовочных таблиц
            }
            else if (tableIndex == 3)
            {
                return "1 семестр";
            }
            else if (tableIndex == 4)
            {
                return "2 семестр";
            }
            else
            {
                return "Не определен";
            }
        }

        private void ProcessWordTable(Word.Table wordTable, DataTable displayTable, string fileName, int tableNumber, string semester)
        {
            try
            {
                // Если таблица пустая (только заголовок), все равно ее обрабатываем
                for (int row = 1; row <= wordTable.Rows.Count; row++)
                {
                    DataRow dataRow = displayTable.NewRow();
                    dataRow["Файл"] = fileName;
                    dataRow["Номер таблицы"] = string.Format("Таблица {0}", tableNumber);
                    dataRow["Семестр"] = semester;

                    // Копируем данные из всех ячеек
                    for (int col = 1; col <= wordTable.Columns.Count && col <= 8; col++)
                    {
                        try
                        {
                            string cellText = wordTable.Cell(row, col).Range.Text;
                            cellText = CleanText(cellText);
                            dataRow[string.Format("Колонка {0}", col)] = cellText;
                        }
                        catch
                        {
                            // Оставляем пустым, если не удалось прочитать ячейку
                        }
                    }

                    displayTable.Rows.Add(dataRow);
                }
            }
            catch (Exception ex)
            {
                AddErrorRow(displayTable, fileName, string.Format("Ошибка при обработке таблицы {0}: {1}", tableNumber, ex.Message));
            }
        }

        private string CleanText(string text)
        {
            if (text == null) return "";

            string cleaned = text
                .Replace("\r", "")
                .Replace("\a", "")
                .Replace("\n", " ")
                .Replace("\v", " ")
                .Replace("\t", " ")
                .Replace("^p", " ")
                .Replace("^t", " ")
                .Replace("\u0007", "")  // Улучшенная обработка специальных символов
                .Replace("\u0008", "")
                .Replace("\\", "")
                .Replace("//", "")
                .Trim();

            // Удаляем лишние пробелы
            while (cleaned.Contains("  "))
                cleaned = cleaned.Replace("  ", " ");

            return cleaned;
        }

        private void AddErrorRow(DataTable displayTable, string fileName, string errorMessage)
        {
            DataRow row = displayTable.NewRow();
            row["Файл"] = fileName;
            row["Номер таблицы"] = "ОШИБКА";
            row["Семестр"] = "";
            row["Колонка 1"] = errorMessage;
            displayTable.Rows.Add(row);
        }

        private void btnExportCSV_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.DataSource == null || dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Нет данных для экспорта", "Информация",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string appPath = Application.StartupPath;
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string csvFilePath = Path.Combine(appPath, string.Format("tables_export_{0}.csv", timestamp));

                ExportToCSV((DataTable)dataGridView1.DataSource, csvFilePath);

                MessageBox.Show(string.Format("Данные успешно экспортированы в CSV файл:\n{0}", csvFilePath),
                              "Экспорт завершен", MessageBoxButtons.OK, MessageBoxIcon.Information);

                UpdateStatus(string.Format("Данные экспортированы: {0}", Path.GetFileName(csvFilePath)));
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при экспорте в CSV: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToCSV(DataTable dataTable, string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                List<string> headers = new List<string>();
                foreach (DataColumn column in dataTable.Columns)
                {
                    headers.Add(column.ColumnName);
                }
                writer.WriteLine(string.Join(";", headers.ToArray()));

                foreach (DataRow row in dataTable.Rows)
                {
                    List<string> values = new List<string>();
                    foreach (object item in row.ItemArray)
                    {
                        string value = item != null ? item.ToString() : "";
                        if (value.Contains(";") || value.Contains("\"") || value.Contains("\n") || value.Contains("\r"))
                        {
                            value = "\"" + value.Replace("\"", "\"\"") + "\"";
                        }
                        values.Add(value);
                    }
                    writer.WriteLine(string.Join(";", values.ToArray()));
                }
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value != null)
            {
                string cellValue = e.Value.ToString();
                if (cellValue == "---")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightGray;
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            UpdateStatus("Готово к работе");
        }

        // НОВЫЙ МЕТОД: Чтение ячейки D2 из CSV файла
        private string ReadCellD2FromCsv(string csvPath)
        {
            try
            {
                using (StreamReader reader = new StreamReader(csvPath, Encoding.UTF8))
                {
                    // Пропускаем первую строку (заголовки)
                    if (!reader.EndOfStream)
                        reader.ReadLine();

                    // Читаем вторую строку (строка 2, индексация с 0)
                    if (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        if (!string.IsNullOrEmpty(line))
                        {
                            string[] values = ParseCsvLine(line, ';');

                            // В CSV файле колонка D соответствует индексу 3 (0-based)
                            // A=0, B=1, C=2, D=3
                            if (values.Length > 3)
                            {
                                return CleanCsvValue(values[3]);

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при чтении ячейки D2: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return "";
        }

        private void ButtonExportWord_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Ищем последний созданный CSV файл
                string latestCsvPath = GetLatestCsvFile();
                if (latestCsvPath == null)
                {
                    MessageBox.Show("CSV файлы не найдены. Сначала экспортируйте данные в CSV.", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 2. Читаем текст из ячейки D2 CSV файла
                string textFromD2 = ReadCellD2FromCsv(latestCsvPath);
                if (string.IsNullOrEmpty(textFromD2))
                {
                    MessageBox.Show("Не удалось прочитать текст из ячейки D2 CSV файла. Убедитесь, что файл содержит данные.", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 3. Читаем CSV и группируем по преподавателям
                var teacherData = ReadCsvDataForTeachers(latestCsvPath);

                if (teacherData.Count == 0)
                {
                    MessageBox.Show("CSV файл пуст или некорректен", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 4. Если несколько преподавателей - показываем диалог выбора
                string selectedTeacher = null;
                if (teacherData.Count > 1)
                {
                    // Создаем список для отображения в форме
                    List<string> teacherNames = new List<string>();
                    foreach (var teacher in teacherData)
                    {
                        // Берем TeacherName, если есть, иначе используем ключ
                        string displayName = !string.IsNullOrEmpty(teacher.Value.TeacherName) ?
                            teacher.Value.TeacherName : teacher.Key;
                        teacherNames.Add(displayName);
                    }

                    // Используем форму выбора с кнопкой "Выбрать"
                    using (var selectForm = new TeacherSelectionFormWithButton(teacherNames, "Выберите преподавателя для экспорта в Word"))
                    {
                        if (selectForm.ShowDialog(this) == DialogResult.OK)
                        {
                            selectedTeacher = selectForm.SelectedTeacher;
                        }
                        else
                        {
                            return; // Пользователь отменил
                        }
                    }
                }
                else
                {
                    var firstTeacher = teacherData.First();
                    selectedTeacher = !string.IsNullOrEmpty(firstTeacher.Value.TeacherName) ?
                        firstTeacher.Value.TeacherName : firstTeacher.Key;
                }

                // 5. Находим данные для выбранного преподавателя
                CsvFileData fileData = null;
                string teacherKey = null;

                foreach (var teacher in teacherData)
                {
                    string displayName = !string.IsNullOrEmpty(teacher.Value.TeacherName) ?
                        teacher.Value.TeacherName : teacher.Key;

                    if (displayName == selectedTeacher)
                    {
                        fileData = teacher.Value;
                        teacherKey = teacher.Key;
                        break;
                    }
                }

                if (fileData == null)
                {
                    MessageBox.Show("Не удалось найти данные для выбранного преподавателя", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 6. Создаем копию шаблона
                string templatePath = Path.Combine(Application.StartupPath, "Шаблон1.docx");
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Шаблон документа не найден", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Создаем имя для нового файла
                string safeTeacherName = selectedTeacher
                    .Replace(" ", "_")
                    .Replace(".", "_")
                    .Replace(",", "_")
                    .Replace(":", "_")
                    .Replace("\\", "_")
                    .Replace("/", "_");

                string newFileName = string.Format("{0}_заполненный.docx", safeTeacherName);
                string outputPath = Path.Combine(Application.StartupPath, newFileName);

                // Удаляем старый файл, если существует
                if (File.Exists(outputPath))
                {
                    try
                    {
                        File.Delete(outputPath);
                    }
                    catch { }
                }

                // Копируем шаблон
                File.Copy(templatePath, outputPath, true);

                // 7. Заполняем документ с передачей текста из D2
                FillWordDocumentWithAllTables(outputPath, fileData, textFromD2);

                MessageBox.Show(string.Format("Документ успешно создан: {0}", newFileName), "Успех",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Открываем документ
                try
                {
                    System.Diagnostics.Process.Start(outputPath);
                }
                catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при экспорте в Word: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для поиска последнего CSV файла
        private string GetLatestCsvFile()
        {
            try
            {
                string appPath = Application.StartupPath;
                string[] csvFiles = Directory.GetFiles(appPath, "tables_export_*.csv");

                if (csvFiles.Length == 0)
                    return null;

                // Сортируем по дате создания (последний созданный файл)
                var fileInfos = new List<FileInfo>();
                foreach (string file in csvFiles)
                {
                    fileInfos.Add(new FileInfo(file));
                }

                fileInfos.Sort((a, b) => b.CreationTime.CompareTo(a.CreationTime));

                return fileInfos[0].FullName;
            }
            catch
            {
                return null;
            }
        }

        // Класс для хранения данных CSV
        public class CsvFileData
        {
            public string Header { get; set; }
            public string Signature { get; set; }
            public string TeacherName { get; set; }
            public string Department { get; set; }
            public List<string[]> Semester1Data { get; set; }
            public List<string[]> Semester2Data { get; set; }

            public CsvFileData()
            {
                Semester1Data = new List<string[]>();
                Semester2Data = new List<string[]>();
            }
        }

        // Метод для чтения CSV данных с группировкой по преподавателям
        private Dictionary<string, CsvFileData> ReadCsvDataForTeachers(string csvPath)
        {
            var result = new Dictionary<string, CsvFileData>();

            try
            {
                using (StreamReader reader = new StreamReader(csvPath, Encoding.UTF8))
                {
                    // Пропускаем заголовок
                    if (!reader.EndOfStream)
                        reader.ReadLine();

                    CsvFileData currentFileData = null;
                    string currentTeacherKey = null;
                    string currentSemester = "";
                    bool processingTeacher = false;
                    string lastFileName = "";

                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        if (string.IsNullOrEmpty(line) || line.Trim().Length == 0)
                            continue;

                        string[] values = ParseCsvLine(line, ';');
                        if (values.Length < 4)
                            continue;

                        string fileName = values[0].Trim();
                        string tableType = values[1].Trim();
                        string semester = values[2].Trim();

                        // Пропускаем разделители
                        if (fileName == "---" || tableType == "---")
                            continue;

                        // Определяем начало нового преподавателя
                        if (tableType == "Документ" && currentFileData == null)
                        {
                            // Начало первого преподавателя в файле
                            currentTeacherKey = "Преподаватель_1";
                            currentFileData = new CsvFileData();
                            processingTeacher = true;
                            currentSemester = "";
                            lastFileName = fileName;

                            // Вторая строка CSV содержит заголовок в колонка 1 (индекс 3)
                            if (values.Length > 3)
                            {
                                currentFileData.Header = CleanCsvValue(values[3]);
                            }
                        }
                        else if (tableType == "Таблица 1")
                        {
                            // Извлекаем имя преподавателя из таблицы 1
                            if (values.Length > 3)
                            {
                                string teacherName = CleanCsvValue(values[3]);
                                if (!string.IsNullOrEmpty(teacherName) && !teacherName.Contains("Дисциплина"))
                                {
                                    // Нормализуем имя преподавателя (убираем дату рождения и оставляем только ФИО)
                                    string normalizedName = NormalizeTeacherName(teacherName);

                                    // Сохраняем предыдущего преподавателя
                                    if (currentFileData != null && !string.IsNullOrEmpty(currentTeacherKey))
                                    {
                                        result[currentTeacherKey] = currentFileData;
                                    }

                                    // Начинаем нового преподавателя с нормализованным именем
                                    currentTeacherKey = normalizedName;
                                    currentFileData = new CsvFileData();
                                    currentFileData.TeacherName = normalizedName;
                                    processingTeacher = true;
                                    currentSemester = "";
                                    lastFileName = fileName;
                                }
                            }
                        }

                        // Если мы не в процессе обработки преподавателя, пропускаем
                        if (!processingTeacher || currentFileData == null)
                            continue;

                        // Проверяем, не сменился ли файл (новый преподаватель)
                        if (fileName != lastFileName && !string.IsNullOrEmpty(fileName) && fileName != "---")
                        {
                            // Сохраняем предыдущего преподавателя
                            if (currentFileData != null && !string.IsNullOrEmpty(currentTeacherKey))
                            {
                                result[currentTeacherKey] = currentFileData;
                            }

                            // Начинаем нового преподавателя
                            currentTeacherKey = string.Format("Преподаватель_{0}", result.Count + 1);
                            currentFileData = new CsvFileData();
                            processingTeacher = true;
                            currentSemester = "";
                            lastFileName = fileName;
                        }

                        // Обновляем текущий семестр
                        if (!string.IsNullOrEmpty(semester) &&
                            (semester == "1 семестр" || semester == "2 семестр"))
                        {
                            currentSemester = semester;
                        }

                        // Обработка разных типов данных
                        if (tableType == "Документ" && string.IsNullOrEmpty(currentFileData.Header))
                        {
                            if (values.Length > 3)
                            {
                                currentFileData.Header = CleanCsvValue(values[3]);
                            }
                        }
                        else if (tableType == "Таблица 2")
                        {
                            // Таблица 2 содержит кафедру
                            if (values.Length > 3)
                            {
                                string department = CleanCsvValue(values[3]);
                                if (!string.IsNullOrEmpty(department) && department.Contains("Кафедра:"))
                                {
                                    currentFileData.Department = department;
                                }
                            }
                        }
                        else if (tableType == "Подпись")
                        {
                            if (values.Length > 3)
                            {
                                string signature = CleanCsvValue(values[3]);
                                currentFileData.Signature = signature;

                                // Если еще не установили имя преподавателя, берем из подписи
                                if (string.IsNullOrEmpty(currentFileData.TeacherName))
                                {
                                    // Из подписи "Преподаватель В. М. Алексеев" извлекаем имя
                                    if (signature.Contains("Преподаватель"))
                                    {
                                        string namePart = signature.Replace("Преподаватель", "").Trim();
                                        currentFileData.TeacherName = namePart;
                                    }
                                }
                            }
                        }
                        else if (tableType.StartsWith("Таблица"))
                        {
                            // Добавляем данные строки таблицы
                            if (values.Length > 3 && !string.IsNullOrEmpty(values[3]))
                            {
                                string col1 = CleanCsvValue(values[3]);

                                // Проверяем, не является ли строка заголовком или итогом
                                if (!col1.Contains("Дисциплина") &&
                                    !col1.Contains("Всего за семестр") &&
                                    !col1.Contains("Всего за учебный год") &&
                                    !col1.Contains("---") &&
                                    !string.IsNullOrEmpty(col1))
                                {
                                    var rowData = new string[8];
                                    for (int i = 0; i < 8 && i + 3 < values.Length; i++)
                                    {
                                        rowData[i] = CleanCsvValue(values[i + 3]);
                                    }

                                    // Добавляем данные в соответствующий семестр
                                    if (currentSemester == "1 семестр")
                                        currentFileData.Semester1Data.Add(rowData);
                                    else if (currentSemester == "2 семестр")
                                        currentFileData.Semester2Data.Add(rowData);
                                }
                            }
                        }
                    }

                    // Добавляем последнего преподавателя
                    if (currentFileData != null && !string.IsNullOrEmpty(currentTeacherKey))
                    {
                        result[currentTeacherKey] = currentFileData;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при чтении CSV файла: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        // Нормализация имени преподавателя (убираем дату рождения, оставляем только ФИО)
        private string NormalizeTeacherName(string teacherName)
        {
            if (string.IsNullOrEmpty(teacherName))
                return teacherName;

            // Разделяем по запятой
            string[] parts = teacherName.Split(',');
            if (parts.Length > 0)
            {
                // Берем первую часть до запятой (ФИО)
                string namePart = parts[0].Trim();

                // Убираем возможную дату рождения в скобках
                int bracketIndex = namePart.IndexOf('(');
                if (bracketIndex > 0)
                {
                    namePart = namePart.Substring(0, bracketIndex).Trim();
                }

                // Убираем лишние пробелы
                namePart = namePart.Replace("  ", " ").Trim();

                return namePart;
            }

            return teacherName.Trim();
        }

        // Парсинг CSV строки с учетом кавычек
        private string[] ParseCsvLine(string line, char delimiter)
        {
            var result = new List<string>();
            bool inQuotes = false;
            StringBuilder currentValue = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // Двойные кавычки внутри кавычек
                        currentValue.Append('"');
                        i++; // Пропускаем следующую кавычку
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == delimiter && !inQuotes)
                {
                    result.Add(currentValue.ToString());
                    currentValue = new StringBuilder();
                }
                else
                {
                    currentValue.Append(c);
                }
            }

            result.Add(currentValue.ToString());
            return result.ToArray();
        }

        // Очистка значения CSV
        private string CleanCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            value = value.Trim();

            // Удаляем кавычки если они есть в начале и конце
            if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"')
            {
                value = value.Substring(1, value.Length - 2);
            }

            // Заменяем двойные кавычки на одинарные
            value = value.Replace("\"\"", "\"");

            return value.Trim();
        }



        // НОВЫЙ МЕТОД: Альтернативный метод поиска TEXT10
        private bool FindAndReplaceText10Alternative(Word.Document doc, string replacementText)
        {
            try
            {
                bool found = false;

                // Метод 2: Поиск по всему тексту документа
                string fullText = doc.Content.Text;
                int text10Index = fullText.IndexOf("TEXT10", StringComparison.Ordinal);

                if (text10Index >= 0)
                {
                    // Если нашли в тексте, пробуем заменить через более точный поиск
                    Word.Range searchRange = doc.Range(0, doc.Content.End);

                    while (searchRange.Find.Execute("TEXT10", Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing))
                    {
                        searchRange.Text = replacementText;
                        found = true;
                    }

                    Marshal.ReleaseComObject(searchRange);
                }



                return found;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при альтернативном поиске TEXT10: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        // ИЗМЕНЕННЫЙ МЕТОД: Добавлен параметр textFromD2 для замены TEXT10
        private void FillWordDocumentWithAllTables(string filePath, CsvFileData fileData, string textFromD2)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                object fileName = filePath;
                object readOnly = false;
                object isVisible = false;
                object missing = Type.Missing;

                doc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly,
                                             ref missing, ref missing, ref missing,
                                             ref missing, ref missing, ref missing,
                                             ref missing, ref missing, ref isVisible,
                                             ref missing, ref missing, ref missing);

                // 1. Заменяем TEXT10 на текст из ячейки D2 CSV файла
                // Пробуем сначала стандартный метод
                bool text10Replaced = FindAndReplaceText10Alternative(doc, textFromD2);

                // Если не удалось, пробуем альтернативный метод


                if (!text10Replaced)
                {
                    MessageBox.Show("Не удалось найти и заменить TEXT10 в документе. Убедитесь, что шаблон содержит текст 'TEXT10'", "Предупреждение",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                // 2. Заменяем TEXT2 на подпись без слова "ПРЕПОДАВАТЕЛЬ"
                if (!string.IsNullOrEmpty(fileData.Signature))
                {
                    string signatureText = fileData.Signature;
                    // Убираем слово "Преподаватель" из подписи
                    signatureText = signatureText.Replace("Преподаватель", "").Trim();

                    // Убираем лишние пробелы, которые могли образоваться
                    while (signatureText.Contains("  "))
                        signatureText = signatureText.Replace("  ", " ");

                    ReplaceTextInDocument(doc, "TEXT2", signatureText);
                }

                // 3. Заполняем все 4 таблицы в документе
                if (doc.Tables.Count >= 4)
                {
                    // Таблица 1: ФИО преподавателя (строка 1, колонка 1)
                    FillTableWithText(doc.Tables[1], !string.IsNullOrEmpty(fileData.TeacherName) ?
                        fileData.TeacherName : "Не указано");

                    // Таблица 2: Кафедра (строка 1, колонка 1)
                    FillTableWithText(doc.Tables[2], !string.IsNullOrEmpty(fileData.Department) ?
                        fileData.Department : "Не указана");

                    // Таблица 3: Данные за 1 семестр
                    FillSemesterTable(doc.Tables[3], fileData.Semester1Data, "1 семестр");

                    // Таблица 4: Данные за 2 семестр
                    FillSemesterTable(doc.Tables[4], fileData.Semester2Data, "2 семестр");
                }
                else if (doc.Tables.Count >= 2)
                {
                    // Если в шаблоне только 2 таблицы, то это таблицы семестров
                    FillSemesterTable(doc.Tables[1], fileData.Semester1Data, "1 семестр");
                    FillSemesterTable(doc.Tables[2], fileData.Semester2Data, "2 семестр");
                }

                // Сохраняем изменения
                object saveFileName = filePath;
                doc.SaveAs(ref saveFileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при заполнении Word документа: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
            finally
            {
                if (doc != null)
                {
                    try
                    {
                        object saveChanges = false;
                        object originalFormat = Word.WdSaveFormat.wdFormatDocumentDefault;
                        object routeDocument = false;

                        doc.Close(ref saveChanges, ref originalFormat, ref routeDocument);
                        Marshal.ReleaseComObject(doc);
                    }
                    catch { }
                }
                if (wordApp != null)
                {
                    try
                    {
                        object saveChangesWord = false;
                        object originalFormatWord = Type.Missing;
                        object routeDocumentWord = Type.Missing;

                        wordApp.Quit(ref saveChangesWord, ref originalFormatWord, ref routeDocumentWord);
                        Marshal.ReleaseComObject(wordApp);
                    }
                    catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Заполнение таблицы одним текстом (для таблиц 1 и 2)
        private void FillTableWithText(Word.Table table, string text)
        {
            try
            {
                if (table == null || string.IsNullOrEmpty(text))
                    return;

                // Заполняем первую ячейку таблицы
                if (table.Rows.Count >= 1 && table.Columns.Count >= 1)
                {
                    Word.Cell cell = table.Cell(1, 1);
                    cell.Range.Text = text;

                    // Устанавливаем нежирное форматирование для соответствия оригиналу
                    cell.Range.Font.Bold = 0;

                    Marshal.ReleaseComObject(cell);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при заполнении таблицы текстом: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Замена текста в документе
        private void ReplaceTextInDocument(Word.Document doc, string findText, string replaceText)
        {
            try
            {
                Word.Range range = doc.Content;

                // Устанавливаем параметры поиска
                range.Find.ClearFormatting();
                range.Find.Text = findText;
                range.Find.Replacement.ClearFormatting();
                range.Find.Replacement.Text = replaceText;

                object replaceAll = Word.WdReplace.wdReplaceAll;
                object missing = Type.Missing;

                // Выполняем замену
                bool found = range.Find.Execute(
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, replaceText,
                    ref replaceAll, ref missing, ref missing, ref missing, ref missing);

                Marshal.ReleaseComObject(range);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при замене текста '{0}': {1}", findText, ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Заполнение таблицы семестра с добавлением строк при необходимости
        private void FillSemesterTable(Word.Table table, List<string[]> data, string semester)
        {
            try
            {
                if (table == null)
                    return;

                // Определяем начальную строку для данных (после заголовка)
                int startRow = 2; // Строка 1 - заголовки, строка 2 - первая строка данных

                // Получаем текущее количество строк в таблице (включая заголовок)
                int existingRows = table.Rows.Count;

                // Рассчитываем, сколько строк нужно добавить
                int rowsNeeded = data.Count;
                int rowsToAdd = rowsNeeded - (existingRows - 1); // -1 потому что первая строка - заголовок

                // Добавляем строки, если нужно больше
                if (rowsToAdd > 0)
                {
                    for (int i = 0; i < rowsToAdd; i++)
                    {
                        // Добавляем новую строку в конец таблицы
                        Word.Row newRow = table.Rows.Add();

                        // Копируем форматирование из предыдущей строки
                        if (existingRows > 1)
                        {
                            // Копируем форматирование из последней строки с данными
                            Word.Row lastRow = table.Rows[existingRows - 1];

                            // Копируем ширину столбцов
                            for (int col = 1; col <= table.Columns.Count; col++)
                            {
                                try
                                {
                                    Word.Cell sourceCell = lastRow.Cells[col];
                                    Word.Cell targetCell = newRow.Cells[col];

                                    // Копируем ширину
                                    targetCell.Width = sourceCell.Width;

                                    // Копируем форматирование
                                    targetCell.Range.Font.Bold = sourceCell.Range.Font.Bold;
                                    targetCell.Range.Font.Size = sourceCell.Range.Font.Size;

                                    Marshal.ReleaseComObject(sourceCell);
                                    Marshal.ReleaseComObject(targetCell);
                                }
                                catch { }
                            }

                            Marshal.ReleaseComObject(lastRow);
                        }

                        Marshal.ReleaseComObject(newRow);
                        existingRows++;
                    }
                }

                // Заполняем данные
                if (data != null && data.Count > 0)
                {
                    for (int i = 0; i < data.Count && i < (existingRows - 1); i++)
                    {
                        string[] rowData = data[i];

                        for (int col = 0; col < Math.Min(rowData.Length, table.Columns.Count); col++)
                        {
                            string cellValue = rowData[col] != null ? rowData[col] : "";

                            // Для пустых значений из CSV оставляем пустую строки
                            if (!string.IsNullOrEmpty(cellValue.Trim()))
                            {
                                table.Cell(startRow + i, col + 1).Range.Text = cellValue;
                            }
                            table.Cell(startRow + i, col + 1).Range.Font.Bold = 0;
                        }
                    }
                }
                else
                {
                    // Если нет данных, добавляем строку "Нет данных"
                    if (existingRows > 1)
                    {
                        table.Cell(startRow, 1).Range.Text = "Нет данных";
                        table.Cell(startRow, 1).Range.Font.Bold = 0;
                    }
                    else
                    {
                        // Добавляем строку, если таблица пустая (только заголовок)
                        Word.Row newRow = table.Rows.Add();
                        table.Cell(startRow, 1).Range.Text = "Нет данных";
                        table.Cell(startRow, 1).Range.Font.Bold = 0;
                        Marshal.ReleaseComObject(newRow);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при заполнении таблицы: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Остальные методы остаются без изменений...

        private void ButtonExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Ищем последний созданный CSV файл
                string latestCsvPath = GetLatestCsvFile();
                if (latestCsvPath == null)
                {
                    MessageBox.Show("CSV файлы не найдены. Сначала экспортируйте данные в CSV.", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 2. Запрашиваем у пользователя период для анализа
                using (var yearForm = new YearSelectionForm())
                {
                    if (yearForm.ShowDialog(this) != DialogResult.OK)
                    {
                        return; // Пользователь отменил
                    }

                    int startYear = yearForm.StartYear;
                    int endYear = yearForm.EndYear;

                    // 3. Читаем CSV файл и фильтруем по году
                    List<string[]> csvData = ReadCsvForExcel(latestCsvPath);
                    List<string[]> filteredData = FilterDataByYear(csvData, startYear, endYear);

                    if (filteredData.Count <= 1) // только заголовок или пусто
                    {
                        // 4. Создаем Excel файл с сообщением об отсутствии данных
                        string excelFileName = string.Format("Excel_отчет_{0}_{1}_{2}.xlsx",
                            startYear, endYear, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                        string excelFilePath = Path.Combine(Application.StartupPath, excelFileName);
                        CreateExcelFileWithNoData(excelFilePath, startYear, endYear);

                        MessageBox.Show(string.Format("Нет данных для периода {0}-{1}. Создан пустой отчет.",
                            startYear, endYear), "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 5. Создаем DataTable для анализа
                    DataTable csvDataTable = ConvertListToDataTable(filteredData);

                    // 6. Извлекаем список преподавателей с проверкой на дубликаты и нормализацией имен
                    Dictionary<string, List<string>> teacherFilesMap = ExtractUniqueTeachersWithFiles(csvDataTable);

                    // Убираем дубликаты преподавателей (нормализуем имена)
                    Dictionary<string, List<string>> normalizedTeacherMap = new Dictionary<string, List<string>>();
                    foreach (var teacher in teacherFilesMap)
                    {
                        string normalizedName = NormalizeTeacherName(teacher.Key);
                        if (!normalizedTeacherMap.ContainsKey(normalizedName))
                        {
                            normalizedTeacherMap[normalizedName] = new List<string>();
                        }
                        // Объединяем файлы для одного преподавателя
                        normalizedTeacherMap[normalizedName].AddRange(teacher.Value);
                    }

                    List<string> teachers = new List<string>(normalizedTeacherMap.Keys);

                    if (teachers.Count == 0)
                    {
                        MessageBox.Show("Не найдено данных о преподавателях", "Информация",
                                      MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 7. Показываем диалог выбора преподавателя с предупреждением о дубликатах
                    string selectedTeacher = ShowTeacherSelectionDialogForExcel(teachers, normalizedTeacherMap);

                    if (string.IsNullOrEmpty(selectedTeacher))
                    {
                        return; // Пользователь отменил
                    }

                    // 8. Получаем список файлов для выбранного преподавателя
                    List<string> teacherFiles = normalizedTeacherMap[selectedTeacher];

                    // 9. Создаем имя для Excel файла
                    string excelFileName2 = string.Format("Excel_отчет_{0}_{1}_{2}.xlsx",
                        startYear, endYear, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                    string excelFilePath2 = Path.Combine(Application.StartupPath, excelFileName2);

                    // 10. Создаем Excel файл с диаграммой для отфильтрованных данных
                    CreateExcelFileWithSingleChart(excelFilePath2, filteredData, startYear, endYear, selectedTeacher, teacherFiles);

                    MessageBox.Show(string.Format("Excel файл успешно создан:\n{0}", excelFileName2), "Успех",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при экспорте в Excel: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для чтения CSV данных для Excel
        private List<string[]> ReadCsvForExcel(string csvPath)
        {
            var result = new List<string[]>();

            using (StreamReader reader = new StreamReader(csvPath, Encoding.UTF8))
            {
                // Читаем заголовки
                string headerLine = reader.ReadLine();
                if (headerLine != null)
                {
                    string[] headers = ParseCsvLine(headerLine, ';');
                    result.Add(headers);
                }

                // Читаем остальные строки
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    if (string.IsNullOrEmpty(line) || line.Trim().Length == 0)
                        continue;

                    string[] values = ParseCsvLine(line, ';');
                    result.Add(values);
                }
            }

            return result;
        }

        // Фильтрация данных по году
        private List<string[]> FilterDataByYear(List<string[]> allData, int startYear, int endYear)
        {
            List<string[]> filteredData = new List<string[]>();

            if (allData.Count == 0)
                return filteredData;

            // Добавляем заголовок
            filteredData.Add(allData[0]);

            for (int i = 1; i < allData.Count; i++)
            {
                string[] row = allData[i];
                if (row.Length < 1 || string.IsNullOrEmpty(row[0]))
                    continue;

                string fileName = row[0];

                // Извлекаем год из имени файла
                int fileYear = ExtractYearFromFileName(fileName);

                // Если год в пределах диапазона, добавляем строку
                if (fileYear >= startYear && fileYear <= endYear)
                {
                    filteredData.Add(row);
                }
            }

            return filteredData;
        }

        // Извлечение года из имени файла
        private int ExtractYearFromFileName(string fileName)
        {
            try
            {
                // Пробуем извлечь год из начала имени файла (например: "2024_Алексеев_В_М_II.docx")
                if (!string.IsNullOrEmpty(fileName))
                {
                    // Ищем первые 4 цифры в имени файла
                    string yearString = "";
                    foreach (char c in fileName)
                    {
                        if (char.IsDigit(c) && yearString.Length < 4)
                        {
                            yearString += c;
                        }
                        else if (yearString.Length == 4)
                        {
                            break;
                        }
                        else if (!char.IsDigit(c) && yearString.Length > 0)
                        {
                            break;
                        }
                    }

                    if (yearString.Length == 4)
                    {
                        return int.Parse(yearString);
                    }

                    // Если не нашли в начале, ищем год в содержимом строки
                    // (например, в заголовке документа)
                    System.Text.RegularExpressions.Regex yearRegex = new System.Text.RegularExpressions.Regex(@"\b(19|20)\d{2}\b");
                    foreach (System.Text.RegularExpressions.Match match in yearRegex.Matches(fileName))
                    {
                        if (match.Success)
                        {
                            return int.Parse(match.Value);
                        }
                    }
                }
            }
            catch
            {
                // В случае ошибки возвращаем 0
            }

            return 0;
        }

        // Создание Excel файла без данных (только сообщение)
        private void CreateExcelFileWithNoData(string filePath, int startYear, int endYear)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Создаем приложение Excel
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // Создаем новую рабочую книгу
                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.Worksheets[1] as Excel.Worksheet;
                worksheet.Name = "Отчет";

                // Добавляем сообщение об отсутствии данных
                worksheet.Cells[1, 1] = "Отчет по почасовой нагрузке";
                worksheet.Cells[2, 1] = string.Format("Период: {0} - {1} гг.", startYear, endYear);
                worksheet.Cells[4, 1] = "Нет данных для выбранного периода";

                // Форматируем сообщение
                Excel.Range titleRange = worksheet.Range["A1", "A1"];
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 14;

                Excel.Range periodRange = worksheet.Range["A2", "A2"];
                periodRange.Font.Bold = true;
                periodRange.Font.Size = 12;

                Excel.Range messageRange = worksheet.Range["A4", "A4"];
                messageRange.Font.Bold = true;
                messageRange.Font.Size = 12;
                messageRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                // Автоподбор ширины столбцов
                worksheet.Columns.AutoFit();

                // Сохраняем файл
                workbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при создании Excel файла: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Извлечение уникальных преподавателей с файлами (с проверкой на дубликаты)
        private Dictionary<string, List<string>> ExtractUniqueTeachersWithFiles(DataTable csvData)
        {
            Dictionary<string, List<string>> teacherFilesMap = new Dictionary<string, List<string>>();

            foreach (DataRow row in csvData.Rows)
            {
                string tableType = row["Номер таблицы"].ToString();
                if (tableType == "Таблица 1")
                {
                    string teacherName = row["Колонка 1"].ToString();
                    string fileName = row["Файл"].ToString();

                    if (!string.IsNullOrEmpty(teacherName) &&
                        !teacherName.Contains("Дисциплина") &&
                        !teacherName.Contains("Всего"))
                    {
                        teacherName = teacherName.Trim();

                        if (!teacherFilesMap.ContainsKey(teacherName))
                        {
                            teacherFilesMap[teacherName] = new List<string>();
                        }

                        if (!teacherFilesMap[teacherName].Contains(fileName))
                        {
                            teacherFilesMap[teacherName].Add(fileName);
                        }
                    }
                }
            }

            return teacherFilesMap;
        }

        // Диалог выбора преподавателя с предупреждением о дубликатах (упрощенная версия)
        private string ShowTeacherSelectionDialogForExcel(List<string> teachers, Dictionary<string, List<string>> teacherFilesMap)
        {
            using (Form selectForm = new Form())
            {
                selectForm.Text = "Выберите преподавателя для диаграммы";
                selectForm.Size = new Size(500, 400);
                selectForm.StartPosition = FormStartPosition.CenterParent;
                selectForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                selectForm.MaximizeBox = false;
                selectForm.MinimizeBox = false;

                // Метка с инструкцией
                Label lblInstruction = new Label();
                lblInstruction.Text = "Выберите преподавателя из списка:";
                lblInstruction.Location = new Point(10, 10);
                lblInstruction.Size = new Size(480, 20);
                lblInstruction.Font = new Font(lblInstruction.Font, FontStyle.Bold);

                // Метка с информацией о дубликатах
                Label lblDuplicateInfo = new Label();
                lblDuplicateInfo.Text = "Преподаватели с несколькими файлами отмечены звездочкой (*)";
                lblDuplicateInfo.Location = new Point(10, 35);
                lblDuplicateInfo.Size = new Size(480, 20);
                lblDuplicateInfo.Font = new Font(lblDuplicateInfo.Font.FontFamily, 9);
                lblDuplicateInfo.ForeColor = Color.DarkRed;

                // ListBox для отображения преподавателей
                ListBox listBox = new ListBox();
                listBox.Location = new Point(10, 60);
                listBox.Size = new Size(480, 250);
                listBox.SelectionMode = SelectionMode.One;

                // Заполняем ListBox с информацией о дубликатах
                foreach (string teacher in teachers)
                {
                    List<string> files = teacherFilesMap[teacher];
                    string displayText = teacher;
                    if (files.Count > 1)
                    {
                        displayText = teacher + " (*) - " + files.Count + " файла(ов)";
                    }
                    listBox.Items.Add(displayText);
                }

                // Кнопка Выбрать
                Button btnSelect = new Button();
                btnSelect.Text = "Выбрать";
                btnSelect.Location = new Point(150, 320);
                btnSelect.Size = new Size(80, 30);
                btnSelect.Click += (s, e) =>
                {
                    if (listBox.SelectedIndex >= 0)
                    {
                        string selectedDisplayText = listBox.SelectedItem.ToString();
                        // Убираем звездочку и информацию о файлах, если они есть
                        string selectedTeacher = selectedDisplayText;
                        if (selectedTeacher.Contains(" (*) - "))
                        {
                            selectedTeacher = selectedTeacher.Substring(0, selectedTeacher.IndexOf(" (*) - "));
                        }
                        else if (selectedTeacher.Contains(" (*)"))
                        {
                            selectedTeacher = selectedTeacher.Replace(" (*)", "");
                        }
                        selectForm.Tag = selectedTeacher;
                        selectForm.DialogResult = DialogResult.OK;
                    }
                    else
                    {
                        MessageBox.Show("Пожалуйста, выберите преподавателя", "Ошибка",
                                      MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                };

                // Кнопка Отмена
                Button btnCancel = new Button();
                btnCancel.Text = "Отмена";
                btnCancel.Location = new Point(250, 320);
                btnCancel.Size = new Size(80, 30);
                btnCancel.Click += (s, e) =>
                {
                    selectForm.DialogResult = DialogResult.Cancel;
                };

                selectForm.Controls.Add(lblInstruction);
                selectForm.Controls.Add(lblDuplicateInfo);
                selectForm.Controls.Add(listBox);
                selectForm.Controls.Add(btnSelect);
                selectForm.Controls.Add(btnCancel);

                if (selectForm.ShowDialog() == DialogResult.OK)
                {
                    return selectForm.Tag as string;
                }
                return null;
            }
        }

        // Создание Excel файла с одной диаграммой (с учетом периода и проверки дубликатов)
        private void CreateExcelFileWithSingleChart(string filePath, List<string[]> data, int startYear, int endYear,
            string selectedTeacher, List<string> teacherFiles)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Создаем приложение Excel
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // Создаем новую рабочую книгу
                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.Worksheets[1] as Excel.Worksheet;
                worksheet.Name = "Данные с диаграммой";

                // Заполняем данные в Excel
                for (int row = 0; row < data.Count; row++)
                {
                    for (int col = 0; col < data[row].Length; col++)
                    {
                        worksheet.Cells[row + 1, col + 1] = data[row][col];
                    }
                }

                // Применяем форматирование
                ApplyExcelFormatting(worksheet, data);

                // Автоподбор ширины столбцов
                worksheet.Columns.AutoFit();

                // Создаем DataTable для анализа
                DataTable csvData = ConvertListToDataTable(data);

                // Рассчитываем все 6 показателей с учетом дубликатов
                Dictionary<string, double> indicators = CalculateAllIndicatorsWithDuplicates(csvData, selectedTeacher, teacherFiles, startYear, endYear);

                // Определяем, где начинать диаграмму (под таблицей)
                int startChartRow = data.Count + 3; // Отступ 3 строки от таблицы

                // Создаем данные для диаграммы
                CreateChartData(worksheet, indicators, startChartRow);

                // Создаем единую диаграмму с указанием периода
                CreateSimpleChartWithDuplicates(worksheet, startChartRow, indicators.Count, selectedTeacher, startYear, endYear, teacherFiles);

                // Сохраняем файл
                workbook.SaveAs(filePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при создании Excel файла: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Конвертация List<string[]> в DataTable
        private DataTable ConvertListToDataTable(List<string[]> data)
        {
            DataTable dt = new DataTable();

            if (data.Count == 0) return dt;

            // Создаем колонки на основе заголовков
            string[] headers = data[0];
            foreach (string header in headers)
            {
                dt.Columns.Add(header);
            }

            // Заполняем строки
            for (int i = 1; i < data.Count; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < Math.Min(headers.Length, data[i].Length); j++)
                {
                    dr[j] = data[i][j] != null ? data[i][j] : "";
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }

        // Расчет всех 6 показателей с учетом дубликатов и суммированием часов (ИСПРАВЛЕННЫЙ РАСЧЕТ ЧАСОВ ПРЕПОДАВАТЕЛЯ)
        private Dictionary<string, double> CalculateAllIndicatorsWithDuplicates(DataTable csvData, string selectedTeacher,
            List<string> teacherFiles, int startYear, int endYear)
        {
            Dictionary<string, double> indicators = new Dictionary<string, double>();

            // 1. Общая сумма часов всех УНИКАЛЬНЫХ преподавателей (без дублирования по файлам)
            double totalHoursAll = 0;
            Dictionary<string, bool> teacherProcessed = new Dictionary<string, bool>();

            foreach (DataRow row in csvData.Rows)
            {
                if (row["Колонка 1"].ToString().Contains("Всего за учебный год"))
                {
                    // Находим преподавателя для этого файла
                    string fileName = row["Файл"].ToString();
                    string teacherName = FindTeacherForFile(fileName, csvData);

                    // Нормализуем имя преподавателя
                    string normalizedTeacherName = NormalizeTeacherName(teacherName);

                    if (!string.IsNullOrEmpty(normalizedTeacherName) && !teacherProcessed.ContainsKey(normalizedTeacherName))
                    {
                        double hours;
                        if (double.TryParse(row["Колонка 4"].ToString(), out hours))
                        {
                            totalHoursAll += hours;
                            teacherProcessed[normalizedTeacherName] = true;
                        }
                    }
                }
            }
            indicators["1. Всего часов"] = totalHoursAll;

            // 2. Сумма часов за первые семестры (только для уникальных преподавателей)
            double totalHoursSemester1 = 0;
            Dictionary<string, bool> teacherProcessedSem1 = new Dictionary<string, bool>();

            foreach (DataRow row in csvData.Rows)
            {
                if (row["Семестр"].ToString() == "1 семестр" &&
                    row["Колонка 1"].ToString().Contains("Всего за семестр"))
                {
                    string fileName = row["Файл"].ToString();
                    string teacherName = FindTeacherForFile(fileName, csvData);

                    // Нормализуем имя преподавателя
                    string normalizedTeacherName = NormalizeTeacherName(teacherName);

                    if (!string.IsNullOrEmpty(normalizedTeacherName) && !teacherProcessedSem1.ContainsKey(normalizedTeacherName))
                    {
                        double hours;
                        if (double.TryParse(row["Колонка 4"].ToString(), out hours))
                        {
                            totalHoursSemester1 += hours;
                            teacherProcessedSem1[normalizedTeacherName] = true;
                        }
                    }
                }
            }
            indicators["2. Часы 1 семестр"] = totalHoursSemester1;

            // 3. Сумма часов за вторые семестры (только для уникальных преподавателей)
            double totalHoursSemester2 = 0;
            Dictionary<string, bool> teacherProcessedSem2 = new Dictionary<string, bool>();

            foreach (DataRow row in csvData.Rows)
            {
                if (row["Семестр"].ToString() == "2 семестр" &&
                    row["Колонка 1"].ToString().Contains("Всего за семестр"))
                {
                    string fileName = row["Файл"].ToString();
                    string teacherName = FindTeacherForFile(fileName, csvData);

                    // Нормализуем имя преподавателя
                    string normalizedTeacherName = NormalizeTeacherName(teacherName);

                    if (!string.IsNullOrEmpty(normalizedTeacherName) && !teacherProcessedSem2.ContainsKey(normalizedTeacherName))
                    {
                        double hours;
                        if (double.TryParse(row["Колонка 4"].ToString(), out hours))
                        {
                            totalHoursSemester2 += hours;
                            teacherProcessedSem2[normalizedTeacherName] = true;
                        }
                    }
                }
            }
            indicators["3. Часы 2 семестр"] = totalHoursSemester2;

            // 4. Количество лекций: (общая сумма часов из п.1) * 60 / 80 (целочисленное деление)
            double lecturesCount = Math.Floor(totalHoursAll * 60 / 80);
            indicators["4. Кол-во лекций"] = lecturesCount;

            // 5. Количество часов для выбранного преподавателя по всем его файлам 
            // ИСПРАВЛЕНИЕ: берем только значение из строки "Всего за учебный год"
            double hoursByGroups = 0;
            if (!string.IsNullOrEmpty(selectedTeacher))
            {
                // Для каждого файла преподавателя суммируем часы из строки "Всего за учебный год"
                foreach (string file in teacherFiles)
                {
                    // Проверяем, попадает ли файл в выбранный период
                    int fileYear = ExtractYearFromFileName(file);
                    if (fileYear >= startYear && fileYear <= endYear)
                    {
                        // Ищем строку "Всего за учебный год" для этого файла
                        foreach (DataRow row in csvData.Rows)
                        {
                            if (row["Файл"].ToString() == file &&
                                row["Колонка 1"].ToString().Contains("Всего за учебный год"))
                            {
                                double hours;
                                if (double.TryParse(row["Колонка 4"].ToString(), out hours))
                                {
                                    hoursByGroups += hours;
                                    break; // Нашли нужную строку, выходим из цикла для этого файла
                                }
                            }
                        }
                    }
                }
            }
            indicators["5. Часы преп-ля"] = hoursByGroups;

            // 6. Количество уникальных видов занятий
            HashSet<string> activityTypes = new HashSet<string>();
            foreach (DataRow row in csvData.Rows)
            {
                string activityType = row["Колонка 3"].ToString();
                if (!string.IsNullOrEmpty(activityType) &&
                    activityType != "Вид занятий" &&
                    !activityType.Contains("Всего"))
                {
                    activityTypes.Add(activityType.Trim());
                }
            }
            indicators["6. Видов занятий"] = activityTypes.Count;

            return indicators;
        }

        // Поиск преподавателя для файла
        private string FindTeacherForFile(string fileName, DataTable csvData)
        {
            foreach (DataRow row in csvData.Rows)
            {
                string currentFileName = row["Файл"].ToString();
                string tableType = row["Номер таблицы"].ToString();

                if (currentFileName == fileName && tableType == "Таблица 1")
                {
                    string teacherName = row["Колонка 1"].ToString();
                    if (!string.IsNullOrEmpty(teacherName) &&
                        !teacherName.Contains("Дисциплина") &&
                        !teacherName.Contains("Всего"))
                    {
                        return teacherName.Trim();
                    }
                }
            }
            return null;
        }

        // Вспомогательный метод для получения короткого имени преподавателя
        private string GetTeacherShortName(string fullName)
        {
            // Пример: "Алексеев Виктор Михайлович, проф., д.т.н." -> "Алексеев"
            string[] parts = fullName.Split(' ');
            if (parts.Length > 0)
            {
                return parts[0];
            }
            return fullName;
        }

        // Создание данных для диаграммы
        private void CreateChartData(Excel.Worksheet worksheet, Dictionary<string, double> indicators, int startRow)
        {
            Excel.Range chartHeaderRange = null;
            Excel.Range chartDataRange = null;
            Excel.Range column1 = null;
            Excel.Range column2 = null;

            try
            {
                int row = startRow;

                // Заголовки
                worksheet.Cells[row, 1] = "Показатель";
                worksheet.Cells[row, 2] = "Значение";

                // Данные
                foreach (var indicator in indicators)
                {
                    row++;
                    worksheet.Cells[row, 1] = indicator.Key;
                    worksheet.Cells[row, 2] = indicator.Value;
                }

                // Форматирование заголовков


                chartHeaderRange.Font.Bold = true;
                chartHeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

                // Форматирование данных


                chartDataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                chartDataRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                // Настройка ширины колонок
                column1 = worksheet.Columns[1] as Excel.Range;
                column1.ColumnWidth = 20;

                column2 = worksheet.Columns[2] as Excel.Range;
                column2.ColumnWidth = 12;
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при создании данных для диаграммы: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                // Освобождаем COM-объекты
                if (column2 != null) Marshal.ReleaseComObject(column2);
                if (column1 != null) Marshal.ReleaseComObject(column1);
                if (chartDataRange != null) Marshal.ReleaseComObject(chartDataRange);
                if (chartHeaderRange != null) Marshal.ReleaseComObject(chartHeaderRange);
            }
        }

        // Создание единой столбчатой диаграммы с 6 показателями (с учетом периода и дубликатов)
        private void CreateSimpleChartWithDuplicates(Excel.Worksheet worksheet, int startRow, int indicatorCount,
            string selectedTeacher, int startYear, int endYear, List<string> teacherFiles)
        {
            try
            {
                // 1. Создаем диапазон для диаграммы
                Excel.Range startCell = worksheet.Cells[startRow, 1] as Excel.Range;
                Excel.Range endCell = worksheet.Cells[startRow + indicatorCount, 2] as Excel.Range;
                Excel.Range dataRange = worksheet.Range[startCell, endCell];

                // 2. Создаем диаграмму (используем альтернативный метод)
                Excel.Shapes shapes = worksheet.Shapes;
                Excel.Shape chartShape = shapes.AddChart(Excel.XlChartType.xlColumnClustered,
                    10, // Left
                    (startRow + indicatorCount + 3) * 15, // Top
                    600, // Width
                    300); // Height

                Excel.Chart excelChart = chartShape.Chart;

                // 3. Устанавливаем данные
                excelChart.SetSourceData(dataRange);

                // 4. Настраиваем заголовок с указанием периода
                excelChart.HasTitle = true;
                string title = string.Format("Статистика почасовой нагрузки ({0}-{1} гг.)", startYear, endYear);
                if (!string.IsNullOrEmpty(selectedTeacher))
                {
                    title += string.Format("\nПреподаватель: {0}", selectedTeacher);

                    // Добавляем информацию о дубликатах
                    if (teacherFiles.Count > 1)
                    {
                        title += string.Format(" (объединено из {0} файлов)", teacherFiles.Count);
                    }
                }
                excelChart.ChartTitle.Text = title;

                // 5. Настраиваем подписи осей (простой способ)
                try
                {
                    // Настройки осей можно добавить при необходимости
                }
                catch
                {
                    // Пропускаем настройку осей, если не получается
                }

                // 6. Настраиваем цвет столбцов
                try
                {
                    Excel.Series series = excelChart.SeriesCollection(1) as Excel.Series;
                    if (series != null)
                    {
                        series.Format.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.SteelBlue);

                        // Включаем подписи значений
                        series.HasDataLabels = true;
                    }
                }
                catch
                {
                    // Пропускаем настройку цвета
                }

                // 7. Освобождаем объекты
                try { Marshal.ReleaseComObject(dataRange); } catch { }
                try { Marshal.ReleaseComObject(chartShape); } catch { }
                try { Marshal.ReleaseComObject(shapes); } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при создании диаграммы: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ApplyExcelFormatting(Excel.Worksheet worksheet, List<string[]> data)
        {
            if (data.Count == 0) return;

            int rowCount = data.Count;
            int colCount = data[0].Length;

            try
            {
                // 1. Форматирование заголовков (первая строка)



                // 2. Форматирование всех данных




                // 3. Форматирование чередующихся строк


                // 4. Настройка ширины столбцов
                for (int col = 1; col <= colCount; col++)
                {
                    Excel.Range column = worksheet.Columns[col] as Excel.Range;
                    column.ColumnWidth = 15; // Базовая ширина

                    // Автоподбор для некоторых колонок
                    if (col == 1) // Файл
                        column.ColumnWidth = 25;
                    else if (col == 2) // Номер таблицы
                        column.ColumnWidth = 18;
                    else if (col == 4) // Колонка 1 (основные данные)
                        column.ColumnWidth = 30;

                    Marshal.ReleaseComObject(column);
                }

                // 5. Заморозка заголовков
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;

                // 6. Автофильтр для заголовков


            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при форматировании Excel: {0}", ex.Message), "Ошибка",
                  MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // МЕТОД ДЛЯ ПОИСКА ПОСЛЕДНЕГО WORD ФАЙЛА
        private string GetLatestWordFile()
        {
            try
            {
                string appPath = Application.StartupPath;
                string[] wordFiles = Directory.GetFiles(appPath, "*_заполненный.docx");

                if (wordFiles.Length == 0)
                    return null;

                // Сортируем по дате создания (последний созданный файл)
                var fileInfos = new List<FileInfo>();
                foreach (string file in wordFiles)
                {
                    fileInfos.Add(new FileInfo(file));
                }

                fileInfos.Sort((a, b) => b.CreationTime.CompareTo(a.CreationTime));

                return fileInfos[0].FullName;
            }
            catch
            {
                return null;
            }
        }

        // МЕТОД ДЛЯ ПОИСКА ПОСЛЕДНЕГО EXCEL ФАЙЛА
        private string GetLatestExcelFile()
        {
            try
            {
                string appPath = Application.StartupPath;
                string[] excelFiles = Directory.GetFiles(appPath, "Excel_отчет_*.xlsx");

                if (excelFiles.Length == 0)
                    return null;

                // Сортируем по дате создания (последний созданный файл)
                var fileInfos = new List<FileInfo>();
                foreach (string file in excelFiles)
                {
                    fileInfos.Add(new FileInfo(file));
                }

                fileInfos.Sort((a, b) => b.CreationTime.CompareTo(a.CreationTime));

                return fileInfos[0].FullName;
            }
            catch
            {
                return null;
            }
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        // ДОПОЛНЕННЫЙ МЕТОД: Открытие последнего созданного CSV файла
        private void BtnOpenCSV_Click(object sender, EventArgs e)
        {
            try
            {
                string latestCsvPath = GetLatestCsvFile();
                if (latestCsvPath == null)
                {
                    MessageBox.Show("CSV файлы не найдены. Сначала экспортируйте данные в CSV.", "Информация",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                System.Diagnostics.Process.Start(latestCsvPath);
                UpdateStatus("Открыт последний CSV файл: " + Path.GetFileName(latestCsvPath));
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при открытии CSV файла: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ДОПОЛНЕННЫЙ МЕТОД: Открытие последнего созданного Word файла
        private void BtnOpenWORD_Click(object sender, EventArgs e)
        {
            try
            {
                string latestWordPath = GetLatestWordFile();
                if (latestWordPath == null)
                {
                    MessageBox.Show("Word файлы не найдены. Сначала экспортируйте данные в Word.", "Информация",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                System.Diagnostics.Process.Start(latestWordPath);
                UpdateStatus("Открыт последний Word файл: " + Path.GetFileName(latestWordPath));
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при открытии Word файла: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ДОПОЛНЕННЫЙ МЕТОД: Открытие последнего созданного Excel файла
        private void BtnOpenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string latestExcelPath = GetLatestExcelFile();
                if (latestExcelPath == null)
                {
                    MessageBox.Show("Excel файлы не найдены. Сначала создайте отчет в Excel.", "Информация",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                System.Diagnostics.Process.Start(latestExcelPath);
                UpdateStatus("Открыт последний Excel файл: " + Path.GetFileName(latestExcelPath));
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при открытии Excel файла: {0}", ex.Message), "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateChart_Click(object sender, EventArgs e)
        {
            try
            {
                // 1. Выбор файлов
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Word Documents (*.docx, *.doc)|*.docx;*.doc";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Выберите файлы для построения диаграммы";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                    return;

                // 2. Выбор периода
                using (YearSelectionForm yearForm = new YearSelectionForm())
                {
                    if (yearForm.ShowDialog() != DialogResult.OK)
                        return;

                    int startYear = yearForm.StartYear;
                    int endYear = yearForm.EndYear;

                    // 3. Выбор типа диаграммы
                    using (ChartTypeForm chartTypeForm = new ChartTypeForm())
                    {
                        if (chartTypeForm.ShowDialog() != DialogResult.OK)
                            return;

                        int chartType = chartTypeForm.SelectedChartType;

                        // 4. Обработка файлов и извлечение данных
                        List<WorkloadRecord> allRecords = new List<WorkloadRecord>();

                        foreach (string filePath in openFileDialog.FileNames)
                        {
                            int fileYear = ExtractYearFromFileName(Path.GetFileName(filePath));

                            // Фильтрация по году
                            if (fileYear >= startYear && fileYear <= endYear)
                            {
                                var records = ExtractWorkloadRecordsFromFile(filePath);
                                allRecords.AddRange(records);
                            }
                        }

                        if (allRecords.Count == 0)
                        {
                            MessageBox.Show("Нет данных для выбранного периода", "Информация",
                                          MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        // 5. Построение диаграммы в зависимости от типа
                        switch (chartType)
                        {
                            case 1: // Общая суммарная нагрузка по преподавателям
                                BuildTotalWorkloadChart(allRecords, startYear, endYear);
                                break;
                            case 2: // Нагрузка в первых семестрах
                                BuildFirstSemesterWorkloadChart(allRecords, startYear, endYear);
                                break;
                            case 3: // Нагрузка во вторых семестрах
                                BuildSecondSemesterWorkloadChart(allRecords, startYear, endYear);
                                break;
                            case 4: // Количество лекций
                                BuildLecturesChart(allRecords, startYear, endYear);
                                break;
                            case 5: // Нагрузка указанного преподавателя по группам
                                BuildTeacherWorkloadByGroupsChart(allRecords, startYear, endYear);
                                break;
                            case 6: // Часы по видам занятий
                                BuildActivityTypesChart(allRecords, startYear, endYear);
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при построении диаграммы: {ex.Message}", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Класс для хранения данных о нагрузке
        private class WorkloadRecord
        {
            public string TeacherName { get; set; }
            public int Year { get; set; }
            public string Semester { get; set; }
            public string ActivityType { get; set; }
            public int Hours { get; set; }
            public bool IsTotalRow { get; set; }
            public string Discipline { get; set; }
            public string FileName { get; set; }
            public string Group { get; set; }
        }

        // Извлечение данных из файла
        private List<WorkloadRecord> ExtractWorkloadRecordsFromFile(string filePath)
        {
            var records = new List<WorkloadRecord>();

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                doc = wordApp.Documents.Open(
                    FileName: filePath,
                    ReadOnly: true,
                    Visible: false
                );

                // Извлечение имени преподавателя
                string teacherName = ExtractTeacherNameFromFile(doc);
                int year = ExtractYearFromFileName(Path.GetFileName(filePath));
                string fileName = Path.GetFileName(filePath);

                // Обработка таблиц семестров (таблицы 3 и 4)
                if (doc.Tables.Count >= 4)
                {
                    // Таблица 3 - первый семестр
                    ProcessSemesterTable(doc.Tables[3], teacherName, year, "1 семестр", records, fileName);

                    // Таблица 4 - второй семестр
                    ProcessSemesterTable(doc.Tables[4], teacherName, year, "2 семестр", records, fileName);
                }
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(false);
                    Marshal.ReleaseComObject(doc);
                }

                if (wordApp != null)
                {
                    wordApp.Quit(false);
                    Marshal.ReleaseComObject(wordApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return records;
        }

        // Извлечение имени преподавателя из файла
        private string ExtractTeacherNameFromFile(Word.Document doc)
        {
            try
            {
                // Ищем в таблице 1
                if (doc.Tables.Count >= 1)
                {
                    Word.Table table = doc.Tables[1];
                    if (table.Rows.Count > 0 && table.Columns.Count > 0)
                    {
                        string cellText = CleanText(table.Cell(1, 1).Range.Text);

                        // Извлекаем только ФИО (до первой запятой)
                        int commaIndex = cellText.IndexOf(',');
                        if (commaIndex > 0)
                        {
                            return cellText.Substring(0, commaIndex).Trim();
                        }

                        return cellText;
                    }
                }
            }
            catch { }

            return "Неизвестный преподаватель";
        }

        // Обработка таблицы семестра с извлечением групп
        private void ProcessSemesterTable(Word.Table table, string teacherName, int year,
                                         string semester, List<WorkloadRecord> records, string fileName)
        {
            try
            {
                // Начинаем со второй строки (первая - заголовок)
                for (int row = 2; row <= table.Rows.Count; row++)
                {
                    try
                    {
                        string discipline = "";
                        string group = "";
                        string activityType = "";
                        string hoursText = "";

                        if (table.Columns.Count >= 1)
                            discipline = CleanText(table.Cell(row, 1).Range.Text);
                        if (table.Columns.Count >= 2)
                            group = CleanText(table.Cell(row, 2).Range.Text);
                        if (table.Columns.Count >= 3)
                            activityType = CleanText(table.Cell(row, 3).Range.Text);
                        if (table.Columns.Count >= 4)
                            hoursText = CleanText(table.Cell(row, 4).Range.Text);

                        // Пропускаем пустые строки и заголовки
                        if (string.IsNullOrEmpty(discipline) ||
                            discipline.Contains("Дисциплина") ||
                            string.IsNullOrEmpty(hoursText))
                            continue;

                        // Проверяем, является ли строка итоговой
                        bool isTotalRow = discipline.Contains("Всего за семестр") ||
                                         discipline.Contains("Всего за учебный год");

                        int hours = 0;
                        if (int.TryParse(hoursText, out hours))
                        {
                            records.Add(new WorkloadRecord
                            {
                                TeacherName = teacherName,
                                Year = year,
                                Semester = semester,
                                Discipline = discipline,
                                Group = group,
                                ActivityType = activityType,
                                Hours = hours,
                                IsTotalRow = isTotalRow,
                                FileName = fileName
                            });
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        // Метод для разбивки текста с группами
        private List<string> SplitGroups(string groupText)
        {
            var result = new List<string>();

            if (string.IsNullOrEmpty(groupText))
                return result;

            // Убираем лишние пробелы и символы
            groupText = groupText.Trim();

            // Разбиваем по разным разделителям
            string[] separators = new[] { ",", ";", "\n", " и ", "/", "\\" };

            // Сначала заменяем все разделители на стандартный
            foreach (var separator in separators)
            {
                groupText = groupText.Replace(separator, "|");
            }

            // Разбиваем по стандартному разделителю
            var parts = groupText.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                string cleanPart = part.Trim();
                if (!string.IsNullOrEmpty(cleanPart))
                    result.Add(cleanPart);
            }

            return result;
        }

        // 1. Диаграмма суммарной нагрузки по преподавателям
        private void BuildTotalWorkloadChart(List<WorkloadRecord> records, int startYear, int endYear)
        {
            CreateChart.Series.Clear();
            CreateChart.Titles.Clear();
            CreateChart.ChartAreas.Clear();

            CreateChart.ChartAreas.Add(new ChartArea());
            CreateChart.Titles.Add($"Суммарная нагрузка преподавателей ({startYear}-{endYear})");

            // Настройки для отображения всех подписей
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            CreateChart.ChartAreas[0].AxisX.Interval = 1; // Показывать каждую подпись
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 8);
            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false; // Убираем отступы
            CreateChart.ChartAreas[0].Position.Auto = false;
            CreateChart.ChartAreas[0].Position.X = 5; // Сдвигаем область диаграммы
            CreateChart.ChartAreas[0].Position.Y = 10;
            CreateChart.ChartAreas[0].Position.Width = 90;
            CreateChart.ChartAreas[0].Position.Height = 80;

            var series = new Series
            {
                Name = "Нагрузка",
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true,
                LabelAngle = -90,
                Font = new Font("Arial", 8)
            };

            // Группируем по преподавателям и суммируем часы
            var data = records
                .Where(r => r.IsTotalRow && r.Discipline.Contains("Всего за учебный год"))
                .GroupBy(r => r.TeacherName)
                .Select(g => new
                {
                    Teacher = g.Key,
                    TotalHours = g.Sum(r => r.Hours)
                })
                .OrderBy(x => x.Teacher)
                .ToList();

            foreach (var item in data)
            {
                var point = series.Points.AddXY(item.Teacher, item.TotalHours);

            }

            CreateChart.Series.Add(series);
            CreateChart.ChartAreas[0].AxisX.Title = "Преподаватели";
            CreateChart.ChartAreas[0].AxisY.Title = "Часы";

            // Автоматическая настройка масштаба оси X
            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].AxisX.IsLabelAutoFit = true;
            CreateChart.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None;
        }

        // 2. Диаграмма нагрузки в первых семестрах
        private void BuildFirstSemesterWorkloadChart(List<WorkloadRecord> records, int startYear, int endYear)
        {
            CreateChart.Series.Clear();
            CreateChart.Titles.Clear();
            CreateChart.ChartAreas.Clear();

            CreateChart.ChartAreas.Add(new ChartArea());
            CreateChart.Titles.Add($"Нагрузка в первых семестрах ({startYear}-{endYear})");

            // Настройки для отображения всех подписей
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            CreateChart.ChartAreas[0].AxisX.Interval = 1;
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 8);
            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].Position.Auto = false;
            CreateChart.ChartAreas[0].Position.X = 5;
            CreateChart.ChartAreas[0].Position.Y = 10;
            CreateChart.ChartAreas[0].Position.Width = 90;
            CreateChart.ChartAreas[0].Position.Height = 80;

            var series = new Series
            {
                Name = "Нагрузка",
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true,
                LabelAngle = -90,
                Font = new Font("Arial", 8)
            };

            var data = records
                .Where(r => r.IsTotalRow && r.Semester == "1 семестр" && r.Discipline.Contains("Всего за семестр"))
                .GroupBy(r => r.TeacherName)
                .Select(g => new
                {
                    Teacher = g.Key,
                    TotalHours = g.Sum(r => r.Hours)
                })
                .OrderBy(x => x.Teacher)
                .ToList();

            foreach (var item in data)
            {
                var point = series.Points.AddXY(item.Teacher, item.TotalHours);

            }

            CreateChart.Series.Add(series);
            CreateChart.ChartAreas[0].AxisX.Title = "Преподаватели";
            CreateChart.ChartAreas[0].AxisY.Title = "Часы";

            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].AxisX.IsLabelAutoFit = true;
            CreateChart.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None;
        }

        // 3. Диаграмма нагрузки во вторых семестрах
        private void BuildSecondSemesterWorkloadChart(List<WorkloadRecord> records, int startYear, int endYear)
        {
            CreateChart.Series.Clear();
            CreateChart.Titles.Clear();
            CreateChart.ChartAreas.Clear();

            CreateChart.ChartAreas.Add(new ChartArea());
            CreateChart.Titles.Add($"Нагрузка во вторых семестрах ({startYear}-{endYear})");

            // Настройки для отображения всех подписей
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            CreateChart.ChartAreas[0].AxisX.Interval = 1;
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 8);
            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].Position.Auto = false;
            CreateChart.ChartAreas[0].Position.X = 5;
            CreateChart.ChartAreas[0].Position.Y = 10;
            CreateChart.ChartAreas[0].Position.Width = 90;
            CreateChart.ChartAreas[0].Position.Height = 80;

            var series = new Series
            {
                Name = "Нагрузка",
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true,
                LabelAngle = -90,
                Font = new Font("Arial", 8)
            };

            var data = records
                .Where(r => r.IsTotalRow && r.Semester == "2 семестр" && r.Discipline.Contains("Всего за семестр"))
                .GroupBy(r => r.TeacherName)
                .Select(g => new
                {
                    Teacher = g.Key,
                    TotalHours = g.Sum(r => r.Hours)
                })
                .OrderBy(x => x.Teacher)
                .ToList();

            foreach (var item in data)
            {
                var point = series.Points.AddXY(item.Teacher, item.TotalHours);

            }

            CreateChart.Series.Add(series);
            CreateChart.ChartAreas[0].AxisX.Title = "Преподаватели";
            CreateChart.ChartAreas[0].AxisY.Title = "Часы";

            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].AxisX.IsLabelAutoFit = true;
            CreateChart.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None;
        }

        // 4. Диаграмма количества лекций
        private void BuildLecturesChart(List<WorkloadRecord> records, int startYear, int endYear)
        {
            CreateChart.Series.Clear();
            CreateChart.Titles.Clear();
            CreateChart.ChartAreas.Clear();

            CreateChart.ChartAreas.Add(new ChartArea());
            CreateChart.Titles.Add($"Количество часов лекций ({startYear}-{endYear})");

            // Настройки для отображения всех подписей
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            CreateChart.ChartAreas[0].AxisX.Interval = 1;
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 8);
            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].Position.Auto = false;
            CreateChart.ChartAreas[0].Position.X = 5;
            CreateChart.ChartAreas[0].Position.Y = 10;
            CreateChart.ChartAreas[0].Position.Width = 90;
            CreateChart.ChartAreas[0].Position.Height = 80;

            var series = new Series
            {
                Name = "Лекции",
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true,
                LabelAngle = -90,
                Font = new Font("Arial", 8)
            };

            var data = records
                .Where(r => !r.IsTotalRow && r.ActivityType.ToLower().Contains("лекции"))
                .GroupBy(r => r.TeacherName)
                .Select(g => new
                {
                    Teacher = g.Key,
                    LectureHours = g.Sum(r => r.Hours)
                })
                .OrderBy(x => x.Teacher)
                .ToList();

            foreach (var item in data)
            {
                var point = series.Points.AddXY(item.Teacher, item.LectureHours);

            }

            CreateChart.Series.Add(series);
            CreateChart.ChartAreas[0].AxisX.Title = "Преподаватели";
            CreateChart.ChartAreas[0].AxisY.Title = "Часы лекций";

            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].AxisX.IsLabelAutoFit = true;
            CreateChart.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None;
        }

        // 5. Диаграмма нагрузки преподавателя по группам (ИСПРАВЛЕННЫЙ МЕТОД)
        private void BuildTeacherWorkloadByGroupsChart(List<WorkloadRecord> records, int startYear, int endYear)
        {
            // Получаем список уникальных преподавателей
            var teachers = records
                .Select(r => r.TeacherName)
                .Distinct()
                .OrderBy(t => t)
                .ToList();

            if (teachers.Count == 0)
            {
                MessageBox.Show("Нет данных о преподавателях", "Информация",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Выбор преподавателя
            using (var teacherForm = new TeacherSelectionFormWithButton(teachers, "Выберите преподавателя"))
            {
                if (teacherForm.ShowDialog() != DialogResult.OK)
                    return;

                string selectedTeacher = teacherForm.SelectedTeacher;

                // Собираем все записи для выбранного преподавателя (не итоговые и с указанной группой)
                var teacherRecords = records.Where(r =>
                    r.TeacherName == selectedTeacher &&
                    !r.IsTotalRow &&
                    !string.IsNullOrEmpty(r.Group)).ToList();

                if (teacherRecords.Count == 0)
                {
                    MessageBox.Show($"Нет данных о группах для преподавателя {selectedTeacher}", "Информация",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Словарь для суммирования часов по группам
                Dictionary<string, int> groupHours = new Dictionary<string, int>();

                foreach (var record in teacherRecords)
                {
                    // Разбиваем группы (могут быть несколько через разделители)
                    var groups = SplitGroups(record.Group);

                    foreach (string group in groups)
                    {
                        string cleanGroup = group.Trim();
                        if (string.IsNullOrEmpty(cleanGroup))
                            continue;

                        // Суммируем часы для каждой группы
                        if (groupHours.ContainsKey(cleanGroup))
                            groupHours[cleanGroup] += record.Hours;
                        else
                            groupHours[cleanGroup] = record.Hours;
                    }
                }

                CreateChart.Series.Clear();
                CreateChart.Titles.Clear();
                CreateChart.ChartAreas.Clear();

                CreateChart.ChartAreas.Add(new ChartArea());
                CreateChart.Titles.Add($"Нагрузка преподавателя {selectedTeacher} по группам ({startYear}-{endYear})");

                // Настройки для отображения всех подписей
                CreateChart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
                CreateChart.ChartAreas[0].AxisX.Interval = 1;
                CreateChart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 8);
                CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
                CreateChart.ChartAreas[0].Position.Auto = false;
                CreateChart.ChartAreas[0].Position.X = 5;
                CreateChart.ChartAreas[0].Position.Y = 10;
                CreateChart.ChartAreas[0].Position.Width = 90;
                CreateChart.ChartAreas[0].Position.Height = 80;

                var series = new Series
                {
                    Name = "Часы",
                    ChartType = SeriesChartType.Column,
                    IsValueShownAsLabel = true,
                    LabelAngle = -90,
                    Font = new Font("Arial", 8)
                };

                // Сортируем группы по алфавиту
                var sortedGroups = groupHours.OrderBy(g => g.Key).ToList();

                foreach (var group in sortedGroups)
                {
                    var point = series.Points.AddXY(group.Key, group.Value);
                   
                }

                CreateChart.Series.Add(series);
                CreateChart.ChartAreas[0].AxisX.Title = "Группы";
                CreateChart.ChartAreas[0].AxisY.Title = "Часы";

                CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
                CreateChart.ChartAreas[0].AxisX.IsLabelAutoFit = true;
                CreateChart.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None;
            }
        }

        // 6. Диаграмма часов по видам занятий
        private void BuildActivityTypesChart(List<WorkloadRecord> records, int startYear, int endYear)
        {
            CreateChart.Series.Clear();
            CreateChart.Titles.Clear();
            CreateChart.ChartAreas.Clear();

            CreateChart.ChartAreas.Add(new ChartArea());
            CreateChart.Titles.Add($"Часы по видам занятий ({startYear}-{endYear})");

            // Настройки для отображения всех подписей
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            CreateChart.ChartAreas[0].AxisX.Interval = 1;
            CreateChart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 8);
            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].Position.Auto = false;
            CreateChart.ChartAreas[0].Position.X = 5;
            CreateChart.ChartAreas[0].Position.Y = 10;
            CreateChart.ChartAreas[0].Position.Width = 90;
            CreateChart.ChartAreas[0].Position.Height = 80;

            var series = new Series
            {
                Name = "Часы",
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true,
                LabelAngle = -90,
                Font = new Font("Arial", 8)
            };

            var data = records
                .Where(r => !r.IsTotalRow && !string.IsNullOrEmpty(r.ActivityType))
                .GroupBy(r => r.ActivityType)
                .Select(g => new
                {
                    ActivityType = g.Key,
                    TotalHours = g.Sum(r => r.Hours)
                })
                .OrderBy(x => x.ActivityType)
                .ToList();

            foreach (var item in data)
            {
                var point = series.Points.AddXY(item.ActivityType, item.TotalHours);

            }

            CreateChart.Series.Add(series);
            CreateChart.ChartAreas[0].AxisX.Title = "Вид занятий";
            CreateChart.ChartAreas[0].AxisY.Title = "Часы";

            CreateChart.ChartAreas[0].AxisX.IsMarginVisible = false;
            CreateChart.ChartAreas[0].AxisX.IsLabelAutoFit = true;
            CreateChart.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None;
        }
    }




}



// Вспомогательная форма для выбора преподавателя с кнопкой "Выбрать"
public class TeacherSelectionFormWithButton : Form
{
    private ListBox listBox;
    private Button btnSelect;
    private Button btnCancel;

    public string SelectedTeacher { get; private set; }

    public TeacherSelectionFormWithButton(List<string> teachers, string title)
    {
        InitializeComponents();
        this.Text = title;
        listBox.DataSource = teachers;
    }

    private void InitializeComponents()
    {
        this.Size = new Size(400, 350);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // Создаем метку с инструкцией
        Label lblInstruction = new Label();
        lblInstruction.Text = "Выберите преподавателя из списка и нажмите кнопку \"Выбрать\":";
        lblInstruction.Location = new Point(10, 10);
        lblInstruction.Size = new Size(380, 30);
        lblInstruction.Font = new Font(lblInstruction.Font, FontStyle.Bold);

        listBox = new ListBox();
        listBox.Location = new Point(10, 50);
        listBox.Size = new Size(380, 200);
        listBox.SelectionMode = SelectionMode.One;

        btnSelect = new Button();
        btnSelect.Text = "Выбрать";
        btnSelect.DialogResult = DialogResult.OK;
        btnSelect.Location = new Point(150, 260);
        btnSelect.Size = new Size(80, 30);
        btnSelect.Click += BtnSelect_Click;

        btnCancel = new Button();
        btnCancel.Text = "Отмена";
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(250, 260);
        btnCancel.Size = new Size(80, 30);

        this.Controls.Add(lblInstruction);
        this.Controls.Add(listBox);
        this.Controls.Add(btnSelect);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnSelect;
        this.CancelButton = btnCancel;
    }

    private void BtnSelect_Click(object sender, EventArgs e)
    {
        if (listBox.SelectedItem != null)
        {
            SelectedTeacher = listBox.SelectedItem.ToString();
            this.DialogResult = DialogResult.OK;
        }
        else
        {
            MessageBox.Show("Пожалуйста, выберите преподавателя", "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}

// Форма для выбора периода (года или диапазона лет)
public class YearSelectionForm : Form
{
    private NumericUpDown numStartYear;
    private NumericUpDown numEndYear;
    private Button btnOK;
    private Button btnCancel;
    private Label lblStartYear;
    private Label lblEndYear;
    private CheckBox chkSingleYear;
    private NumericUpDown numSingleYear;

    public int StartYear { get; private set; }
    public int EndYear { get; private set; }

    public YearSelectionForm()
    {
        InitializeComponents();
        this.Text = "Выберите период для анализа";
        this.StartYear = DateTime.Now.Year - 1;
        this.EndYear = DateTime.Now.Year;
    }

    private void InitializeComponents()
    {
        this.Size = new Size(400, 250);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        // CheckBox для выбора одного года
        chkSingleYear = new CheckBox();
        chkSingleYear.Text = "Только один год";
        chkSingleYear.Location = new Point(20, 20);
        chkSingleYear.Size = new Size(150, 25);
        chkSingleYear.CheckedChanged += ChkSingleYear_CheckedChanged;

        // Поле для выбора одного года
        numSingleYear = new NumericUpDown();
        numSingleYear.Location = new Point(180, 20);
        numSingleYear.Size = new Size(80, 25);
        numSingleYear.Minimum = 2000;
        numSingleYear.Maximum = 2100;
        numSingleYear.Value = DateTime.Now.Year;
        numSingleYear.Enabled = false;

        // Метка для начального года
        lblStartYear = new Label();
        lblStartYear.Text = "Начальный год:";
        lblStartYear.Location = new Point(20, 60);
        lblStartYear.Size = new Size(100, 25);

        // Поле для начального года
        numStartYear = new NumericUpDown();
        numStartYear.Location = new Point(130, 60);
        numStartYear.Size = new Size(80, 25);
        numStartYear.Minimum = 2000;
        numStartYear.Maximum = 2100;
        numStartYear.Value = DateTime.Now.Year - 1;

        // Метка для конечного года
        lblEndYear = new Label();
        lblEndYear.Text = "Конечный год:";
        lblEndYear.Location = new Point(20, 100);
        lblEndYear.Size = new Size(100, 25);

        // Поле для конечного года
        numEndYear = new NumericUpDown();
        numEndYear.Location = new Point(130, 100);
        numEndYear.Size = new Size(80, 25);
        numEndYear.Minimum = 2000;
        numEndYear.Maximum = 2100;
        numEndYear.Value = DateTime.Now.Year;

        // Кнопка OK
        btnOK = new Button();
        btnOK.Text = "OK";
        btnOK.DialogResult = DialogResult.OK;
        btnOK.Location = new Point(150, 160);
        btnOK.Size = new Size(80, 30);
        btnOK.Click += BtnOK_Click;

        // Кнопка Отмена
        btnCancel = new Button();
        btnCancel.Text = "Отмена";
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(250, 160);
        btnCancel.Size = new Size(80, 30);

        this.Controls.Add(chkSingleYear);
        this.Controls.Add(numSingleYear);
        this.Controls.Add(lblStartYear);
        this.Controls.Add(numStartYear);
        this.Controls.Add(lblEndYear);
        this.Controls.Add(numEndYear);
        this.Controls.Add(btnOK);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;
    }

    private void ChkSingleYear_CheckedChanged(object sender, EventArgs e)
    {
        bool singleYearMode = chkSingleYear.Checked;

        numSingleYear.Enabled = singleYearMode;
        numStartYear.Enabled = !singleYearMode;
        numEndYear.Enabled = !singleYearMode;
        lblStartYear.Enabled = !singleYearMode;
        lblEndYear.Enabled = !singleYearMode;
    }

    private void BtnOK_Click(object sender, EventArgs e)
    {
        if (chkSingleYear.Checked)
        {
            StartYear = (int)numSingleYear.Value;
            EndYear = (int)numSingleYear.Value;
        }
        else
        {
            StartYear = (int)numStartYear.Value;
            EndYear = (int)numEndYear.Value;

            if (StartYear > EndYear)
            {
                MessageBox.Show("Начальный год не может быть больше конечного", "Ошибка",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}

// Новая форма для выбора типа диаграммы
public class ChartTypeForm : Form
{
    private ComboBox cmbChartType;
    private Button btnOK;
    private Button btnCancel;

    public int SelectedChartType { get; private set; }

    public ChartTypeForm()
    {
        InitializeComponents();
        this.Text = "Выберите тип диаграммы";
    }

    private void InitializeComponents()
    {
        this.Size = new Size(500, 200);
        this.StartPosition = FormStartPosition.CenterParent;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;

        Label label = new Label();
        label.Text = "Выберите тип диаграммы:";
        label.Location = new Point(10, 20);
        label.Size = new Size(200, 25);
        label.Font = new Font(label.Font, FontStyle.Bold);

        cmbChartType = new ComboBox();
        cmbChartType.Location = new Point(10, 50);
        cmbChartType.Size = new Size(460, 25);
        cmbChartType.DropDownStyle = ComboBoxStyle.DropDownList;

        // Добавляем 6 вариантов диаграмм
        cmbChartType.Items.Add("1. Суммарная нагрузка преподавателей за период");
        cmbChartType.Items.Add("2. Нагрузка преподавателей в первых семестрах");
        cmbChartType.Items.Add("3. Нагрузка преподавателей во вторых семестрах");
        cmbChartType.Items.Add("4. Количество часов лекций по преподавателям");
        cmbChartType.Items.Add("5. Нагрузка выбранного преподавателя по группам");
        cmbChartType.Items.Add("6. Часы по видам занятий (все преподаватели)");

        cmbChartType.SelectedIndex = 0;

        btnOK = new Button();
        btnOK.Text = "OK";
        btnOK.DialogResult = DialogResult.OK;
        btnOK.Location = new Point(150, 100);
        btnOK.Size = new Size(80, 30);
        btnOK.Click += BtnOK_Click;

        btnCancel = new Button();
        btnCancel.Text = "Отмена";
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(250, 100);
        btnCancel.Size = new Size(80, 30);

        this.Controls.Add(label);
        this.Controls.Add(cmbChartType);
        this.Controls.Add(btnOK);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;
    }

    private void BtnOK_Click(object sender, EventArgs e)
    {
        SelectedChartType = cmbChartType.SelectedIndex + 1; // +1 потому что индексы с 0, а типы с 1
    }
}