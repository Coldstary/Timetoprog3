using System;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections.Generic;
using Ex = Microsoft.Office.Interop.Excel; // Псевдоним для Excel Interop

namespace ExcelWork20
{
    public partial class Form1 : Form
    {
        private Ex.Application excelApp; // Экземпляр приложения Excel
        private string csvFilePath; // Путь к файлу CSV с данными

        public Form1()
        {
            InitializeComponent();
            // Формируем полный путь к CSV файлу в папке приложения
            csvFilePath = Path.Combine(Application.StartupPath, "данныеЗадачи20.csv");

            string message = "Добро пожаловать в программу ExcelWork20!\n\n" +
                            "Эта программа создает таблицы нагрузки на основе данных CSV.\n\n" +
                            "Краткое пособие:\n" +
                            "1. Нажмите 'Data_entry' для редактирования CSV файла\n" +
                            "(при работе через блокнот учтите что разделителем столбцов является символ ;\n" +
                            "переход на новую строку в файле соответсвует переходу на новую строку в таблице Excel)\n\n" +
                                "2. Нажмите 'Create file' для генерации Excel файла на основе данных из файла CSV\n\n" +
                                "3. Нажмите 'Exit' для завершения работы\n\n" +
                                "Путь к CSV файлу: " + csvFilePath;

                MessageBox.Show(message, "Добро пожаловать в ExcelWork20",
                              MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }
      
        // Обработчик нажатия кнопки "Создать таблицу"
        private void buttonOpen_Click(object sender, EventArgs e)
        {
            try
            {
                // Создаем уникальное имя файла Excel с временной меткой
                string excelFileName = $"таблица нагрузки_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                string excelFilePath = Path.Combine(Application.StartupPath, excelFileName);

                // Проверяем существование CSV файла
                if (!File.Exists(csvFilePath))
                {
                    MessageBox.Show($"CSV файл не найден: {csvFilePath}", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Инициализируем приложение Excel
                excelApp = new Ex.Application();
                // Создаем новую книгу и получаем первый лист
                Ex.Workbook workbook = excelApp.Workbooks.Add();
                Ex.Worksheet worksheet = (Ex.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "ExcelWorkTime20"; // Задаем имя листа

                // Создаем структуру таблицы с заголовками
                CreateExactTableStructure(worksheet);

                // Заполняем таблицу данными из CSV файла
                FillDataFromCsv(worksheet, csvFilePath);

                // Применяем форматирование к таблице
                FormatExcelTableExactly(worksheet);

                // Сохраняем файл и делаем Excel видимым
                workbook.SaveAs(excelFilePath);
                excelApp.Visible = true;
                worksheet.Activate(); // Активируем лист

                // Информируем пользователя об успешном создании
                MessageBox.Show($"Таблица нагрузки успешно создана:\n{excelFilePath}",
                               "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Обработка ошибок с выводом детальной информации
                MessageBox.Show($"Ошибка при создании таблицы:\n{ex.Message}\n{ex.StackTrace}",
                              "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод создания фиксированной структуры таблицы
        private void CreateExactTableStructure(Ex.Worksheet worksheet)
        {
            // Заполняем заголовки в первой строке
            worksheet.Cells[1, 6] = "количество"; // Столбец F
            worksheet.Cells[1, 9] = "Распределение учебной нагрузки(в часах)"; // Столбец I
            worksheet.Cells[1, 23] = "ВCЕГО"; // Столбец W
            worksheet.Cells[1, 24] = "в том числе"; // Столбец X

            // Массив заголовков для второй строки (столбцы A-Z)
            string[] columnHeaders = {
                "N","Наименование дисциплины","Принимаемая нагрузка", "Группы",
                "Семестр", "Студентов", "Потоков", "Групп", "Лекции", "практики",
                "лабораторные", "консультации", "контр работы", "Зачеты", "экзамены",
                "уч Практика", "пр практика", "курсовые", "дипломные", "ГЭК",
                "аспирантура", "вступит экз", "", "очное", "очно-заочное", "заочное"
            };

            // Записываем заголовки во вторую строку
            for (int i = 0; i < columnHeaders.Length; i++)
            {
                worksheet.Cells[2, i + 1] = columnHeaders[i];
            }
            // Переопределяем заголовок первого столбца
            worksheet.Cells[2, 1] = "N";
        }

        // Метод заполнения таблицы данными из CSV файла
        private void FillDataFromCsv(Ex.Worksheet worksheet, string csvFilePath)
        {
            try
            {
                // Чтение всех строк CSV файла с кодировкой Windows-1251
                string[] csvLines = File.ReadAllLines(csvFilePath, Encoding.GetEncoding(1251));

                if (csvLines.Length == 0)
                    return; // Если файл пустой, выходим

                char separator = ';'; // Разделитель в CSV файле
                int excelRow = 3; // Начинаем заполнение с третьей строки Excel

                // Обрабатываем каждую строку CSV файла
                for (int csvRowIndex = 0; csvRowIndex < csvLines.Length; csvRowIndex++)
                {
                    string csvLine = csvLines[csvRowIndex];
                    // Разбиваем строку на элементы по разделителю
                    string[] csvData = csvLine.Split(new char[] { separator }, StringSplitOptions.None);

                    // Записываем данные в Excel, начиная со столбца D (индекс 4)
                    for (int csvColIndex = 0; csvColIndex < csvData.Length; csvColIndex++)
                    {
                        int excelCol = csvColIndex + 4; // Столбец D и далее
                        string value = csvData[csvColIndex];
                        worksheet.Cells[excelRow, excelCol] = value;
                    }

                    excelRow++; // Переходим к следующей строке Excel

                    // Ограничиваем таблицу 8 строками данных
                    if (excelRow > 8)
                        break;
                }

                // Заголовки для строк, которые не заполняются из CSV
                string[] defaultRowTitles = {
                    "Дополнительно",
                    "Итого за 1 семестр",
                    "Геология и геохимия нефти и газа (+)(ОПД)",
                    "Геология и геохимия нефти и газа (+)(ОПД)",
                    "Итого за 2 семестр",
                    "Итого за год"
                };

                // Заполняем служебные ячейки (номера строк и статические заголовки)
                for (int row = 3; row <= 8; row++)
                {
                    // Нумерация строк
                    if (row == 3)
                        worksheet.Cells[row, 1] = 1;
                    else if (row == 5)
                        worksheet.Cells[row, 1] = 1;
                    else if (row == 6)
                        worksheet.Cells[row, 1] = 2;
                    else
                        worksheet.Cells[row, 1] = row - 2;

                    // Заголовки строк
                    if (row - 3 < defaultRowTitles.Length)
                        worksheet.Cells[row, 2] = defaultRowTitles[row - 3];

                    // Дополнительные описания для определенных строк
                    if (row == 5)
                        worksheet.Cells[row, 3] = "Гидрогеология и инженерная геология, ГИДРОГЕОЛОГ, ИНЖЕНЕР-ГЕОЛОГ, очное";
                    else if (row == 6)
                        worksheet.Cells[row, 3] = "Геология, ГЕОЛОГ (не предусмотрено), очное";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении CSV файла:\n{ex.Message}",
                              "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод форматирования таблицы
        private void FormatExcelTableExactly(Ex.Worksheet worksheet)
        {
            try
            {
                // 1. Форматирование первой строки заголовков
                Ex.Range header1 = worksheet.Range["A1:Z1"];
                header1.Font.Bold = false;
                header1.Font.Size = 11;
                header1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                header1.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                header1.RowHeight = 25;

                // 2. Форматирование второй строки (вертикальный текст)
                Ex.Range header2 = worksheet.Range["A2:Z2"];
                header2.Orientation = 90; // Поворот текста на 90 градусов
                header2.Font.Bold = false;
                header2.Font.Size = 9;
                header2.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                header2.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                header2.WrapText = true; // Перенос текста
                header2.RowHeight = 100; // Увеличенная высота для вертикального текста
                header2.Interior.Color = ColorTranslator.ToOle(Color.White); // Белый фон

                // 3. Объединение ячеек и специальное форматирование
                try
                {
                    // Объединение ячейки "ВСЕГО" (W1:W2)
                    Ex.Range rangeW1W2 = worksheet.Range["W1:W2"];
                    rangeW1W2.Merge();
                    rangeW1W2.Orientation = 90;
                    rangeW1W2.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                    rangeW1W2.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                    rangeW1W2.Font.Size = 9;
                    rangeW1W2.WrapText = true;

                    // Объединение ячеек A-E по вертикали
                    for (int col = 1; col <= 5; col++)
                    {
                        Ex.Range cellRange = worksheet.Range[worksheet.Cells[1, col], worksheet.Cells[2, col]];
                        cellRange.Merge();
                        cellRange.Orientation = 0; // Горизонтальный текст
                        cellRange.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                        cellRange.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                        cellRange.Font.Size = 9;
                        cellRange.WrapText = true;

                        // Очистка дублирующихся значений
                        if (col == 1)
                            worksheet.Cells[1, col] = "N";
                        else
                            worksheet.Cells[1, col] = "";
                        worksheet.Cells[2, col] = "";
                    }

                    // Установка правильных заголовков для объединенных ячеек
                    worksheet.Cells[1, 2] = "Наименование дисциплины";
                    worksheet.Cells[1, 3] = "Принимаемая нагрузка";
                    worksheet.Cells[1, 4] = "Группы";
                    worksheet.Cells[1, 5] = "Семестр";

                    // Объединение ячеек для заголовка "количество" (F1:H1)
                    Ex.Range rangeF1H1 = worksheet.Range["F1:H1"];
                    rangeF1H1.Merge();
                    rangeF1H1.Orientation = 0;
                    rangeF1H1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;

                    // Объединение ячеек для заголовка нагрузки (I1:V1)
                    Ex.Range rangeI1V1 = worksheet.Range["I1:V1"];
                    rangeI1V1.Merge();
                    rangeI1V1.Orientation = 0;
                    rangeI1V1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;

                    // Объединение ячеек для заголовка "в том числе" (X1:Z1)
                    Ex.Range rangeX1Z1 = worksheet.Range["X1:Z1"];
                    rangeX1Z1.Merge();
                    rangeX1Z1.Orientation = 0;
                    rangeX1Z1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при объединении: {ex.Message}");
                }

                // 4. Выделение цветом определенных строк
                Color darkGreen = Color.FromArgb(0, 100, 0); // Темно-зеленый цвет
                int darkGreenOle = ColorTranslator.ToOle(darkGreen);

                // Строка 4 (A4:Z4)
                Ex.Range row3 = worksheet.Range["A4:Z4"];
                row3.Interior.Color = darkGreenOle;
                row3.Font.Color = ColorTranslator.ToOle(Color.White);

                // Строка 7 (A7:Z7)
                Ex.Range row6 = worksheet.Range["A7:Z7"];
                row6.Interior.Color = darkGreenOle;
                row6.Font.Color = ColorTranslator.ToOle(Color.White);

                // Строка 8 (A8:Z8)
                Ex.Range row7 = worksheet.Range["A8:Z8"];
                row7.Interior.Color = darkGreenOle;
                row7.Font.Color = ColorTranslator.ToOle(Color.White);

                // 5. Настройка числового формата и выравнивания
                for (int col = 1; col <= 26; col++)
                {
                    Ex.Range columnRange = worksheet.Range[worksheet.Cells[3, col], worksheet.Cells[8, col]];

                    // Для числовых столбцов (A, D-Z)
                    if (col == 1 || (col >= 4 && col <= 26))
                    {
                        columnRange.NumberFormat = "0"; // Целочисленный формат
                        columnRange.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                    }
                    // Для текстовых столбцов (B, C)
                    else if (col == 2 || col == 3)
                    {
                        columnRange.HorizontalAlignment = Ex.XlHAlign.xlHAlignLeft;
                        columnRange.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                    }
                }

                // 6. Добавление границ таблицы
                Ex.Range tableRange = worksheet.Range["A1:Z8"];
                tableRange.Borders.LineStyle = Ex.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Ex.XlBorderWeight.xlThin;

                // 7. Настройка ширины столбцов
                for (int col = 1; col <= 26; col++)
                {
                    Ex.Range column = (Ex.Range)worksheet.Columns[col];

                    if (col == 1 || (col >= 4 && col <= 8))
                    {
                        column.ColumnWidth = 6; // Узкие столбцы для цифр
                    }
                    else if (col == 2)
                    {
                        column.ColumnWidth = 35; // Широкий для названий дисциплин
                    }
                    else if (col == 3)
                    {
                        column.ColumnWidth = 50; // Очень широкий для описаний
                    }
                    else if (col >= 9 && col <= 22)
                    {
                        column.ColumnWidth = 4; // Очень узкие для часов нагрузки
                    }
                    else if (col == 23) // Столбец W (ВСЕГО)
                    {
                        column.ColumnWidth = 8;
                    }
                    else if (col >= 24 && col <= 26) // Столбцы X, Y, Z
                    {
                        column.ColumnWidth = 10;
                    }
                }

                // 8. Настройка высоты строк с данными
                for (int row = 3; row <= 8; row++)
                {
                    Ex.Range rowRange = (Ex.Range)worksheet.Rows[row];
                    rowRange.RowHeight = 25;
                    rowRange.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                }

                // 9. Настройка шрифта для всей таблицы
                Ex.Range allCells = worksheet.Range["A1:Z8"];
                allCells.Font.Name = "Calibri";
                allCells.Font.Size = 11;

                // 10. Фиксация высоты строк заголовков
                ((Ex.Range)worksheet.Rows[1]).RowHeight = 25;
                ((Ex.Range)worksheet.Rows[2]).RowHeight = 100;

                // 11. Изменение цвета текста для определенных ячеек
                Color greenColor = Color.Green;
                int greenOle = ColorTranslator.ToOle(greenColor);

                // Ячейка B3 ("Дополнительно")
                Ex.Range cellB3 = worksheet.Cells[3, 2] as Ex.Range;
                if (cellB3 != null)
                {
                    cellB3.Font.Color = (object)greenOle; // Явное приведение типа
                }

                // Ячейка B5 (первая "Геология и геохимия нефти и газа")
                Ex.Range cellB5 = worksheet.Cells[5, 2] as Ex.Range;
                if (cellB5 != null)
                {
                    cellB5.Font.Color = (object)greenOle;
                }

                // Ячейка B6 (вторая "Геология и геохимия нефти и газа")
                Ex.Range cellB6 = worksheet.Cells[6, 2] as Ex.Range;
                if (cellB6 != null)
                {
                    cellB6.Font.Color = (object)greenOle;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка форматирования: {ex.Message}");
            }
        }

        // Обработчик кнопки "Ввод данных" (открытие CSV файла для редактирования)
        private void Data_entry_Click(object sender, EventArgs e)
        {
            try
            {
                // Если CSV файл не существует, создаем пустой
                if (!File.Exists(csvFilePath))
                {
                    File.WriteAllText(csvFilePath, "", Encoding.GetEncoding(1251));
                }

                // Открываем файл в программе по умолчанию
                Process.Start(csvFilePath);

                // Инструкция для пользователя
                MessageBox.Show($"Файл '{csvFilePath}' открыт для редактирования.\n" +
                               "Внесите изменения и сохраните файл, затем нажмите кнопку 'Создать таблицу'.",
                               "Редактирование CSV", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии файла:\n{ex.Message}",
                              "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Обработчик кнопки "Выход" (заглушка)
        private void exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}