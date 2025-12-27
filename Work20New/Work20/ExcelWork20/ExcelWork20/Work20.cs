using System;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Collections.Generic;
using Ex = Microsoft.Office.Interop.Excel; // Псевдоним для Excel Interop для упрощения использования

namespace ExcelWork20
{
    public partial class Work20 : Form
    {
        private Ex.Application excelApp; // Экземпляр приложения Excel для взаимодействия с ним
        private string csvFilePath; // Путь к файлу CSV с исходными данными

        public Work20()
        {
            InitializeComponent();
            // Формирование полного пути к CSV файлу в папке приложения для хранения данных
            csvFilePath = Path.Combine(Application.StartupPath, "данныеЗадачи20.csv");
        }

        // Обработчик события нажатия кнопки "Создать таблицу"
        private void buttonOpen_Click(object sender, EventArgs e)
        {
            try
            {
                // Создание уникального имени файла Excel с временной меткой для избежания конфликтов
                string excelFileName = $"таблица нагрузки_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                string excelFilePath = Path.Combine(Application.StartupPath, excelFileName);

                // Проверка существования CSV файла для предотвращения ошибок при чтении
                if (!File.Exists(csvFilePath))
                {
                    MessageBox.Show($"CSV файл не найден: {csvFilePath}", "Ошибка",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Инициализация приложения Excel и создание новой рабочей книги
                excelApp = new Ex.Application();
                Ex.Workbook workbook = excelApp.Workbooks.Add();
                Ex.Worksheet worksheet = (Ex.Worksheet)workbook.Worksheets[1];
                worksheet.Name = "ExcelWorkTime20"; // Установка имени листа для идентификации

                // Создание фиксированной структуры таблицы с заголовками и названиями столбцов
                CreateExactTableStructure(worksheet);

                // Заполнение таблицы данными из CSV файла для интеграции информации
                FillDataFromCsv(worksheet, csvFilePath);

                // Применение форматирования к таблице для улучшения читаемости и внешнего вида
                FormatExcelTableExactly(worksheet);

                // Сохранение файла и отображение Excel пользователю
                workbook.SaveAs(excelFilePath);
                excelApp.Visible = true;
                worksheet.Activate(); // Активация листа для немедленного просмотра

                // Информирование пользователя об успешном создании таблицы
                MessageBox.Show($"Таблица нагрузки успешно создана:\n{excelFilePath}",
                               "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // Обработка исключений с выводом детальной информации об ошибке
                MessageBox.Show($"Ошибка при создании таблицы:\n{ex.Message}\n{ex.StackTrace}",
                              "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод создания точной структуры таблицы с заголовками и названиями столбцов
        private void CreateExactTableStructure(Ex.Worksheet worksheet)
        {
            // Заполнение заголовков в первой строке для основных разделов таблицы
            worksheet.Cells[1, 6] = "количество"; // Столбец F - количество студентов/групп
            worksheet.Cells[1, 9] = "Распределение учебной нагрузки(в часах)"; // Столбец I - нагрузка по видам
            worksheet.Cells[1, 23] = "ВCЕГО"; // Столбец W - итоговая нагрузка
            worksheet.Cells[1, 24] = "в том числе"; // Столбец X - детализация по формам обучения

            // Массив заголовков для второй строки - названия каждого столбца (A-Z)
            string[] columnHeaders = {
                "N","Наименование дисциплины","Принимаемая нагрузка", "Группы",
                "Семестр", "Студентов", "Потоков", "Групп", "Лекции", "практики",
                "лабораторные", "консультации", "контр работы", "Зачеты", "экзамены",
                "уч Практика", "пр практика", "курсовые", "дипломные", "ГЭК",
                "аспирантура", "вступит экз", "", "очное", "очно-заочное", "заочное"
            };

            // Запись заголовков во вторую строку каждого столбца
            for (int i = 0; i < columnHeaders.Length; i++)
            {
                worksheet.Cells[2, i + 1] = columnHeaders[i];
            }
            // Переопределение заголовка первого столбца для стандартизации
            worksheet.Cells[2, 1] = "N";
        }

        // Метод заполнения таблицы данными из CSV файла
        private void FillDataFromCsv(Ex.Worksheet worksheet, string csvFilePath)
        {
            try
            {
                // Чтение всех строк CSV файла с указанием кодировки Windows-1251 для корректного отображения кириллицы
                string[] csvLines = File.ReadAllLines(csvFilePath, Encoding.GetEncoding(1251));

                if (csvLines.Length == 0)
                    return; // Если файл пустой, прекращаем обработку

                char separator = ';'; // Разделитель полей в CSV файле (точка с запятой)
                int excelRow = 3; // Начальная строка для заполнения данных в Excel (после заголовков)

                // Обработка каждой строки CSV файла для переноса данных в таблицу Excel
                for (int csvRowIndex = 0; csvRowIndex < csvLines.Length; csvRowIndex++)
                {
                    string csvLine = csvLines[csvRowIndex];
                    // Разделение строки CSV на отдельные элементы данных по разделителю
                    string[] csvData = csvLine.Split(new char[] { separator }, StringSplitOptions.None);

                    // Запись данных в ячейки Excel, начиная со столбца D (4-й столбец)
                    for (int csvColIndex = 0; csvColIndex < csvData.Length; csvColIndex++)
                    {
                        int excelCol = csvColIndex + 4; // Смещение для записи в столбцы D и далее
                        string value = csvData[csvColIndex];
                        worksheet.Cells[excelRow, excelCol] = value;
                    }

                    excelRow++; // Переход к следующей строке Excel для следующей записи

                    // Ограничение таблицы 8 строками данных для сохранения структуры
                    if (excelRow > 8)
                        break;
                }

                // Массив стандартных заголовков для строк, которые не заполняются из CSV
                string[] defaultRowTitles = {
                    "Дополнительно",
                    "Итого за 1 семестр",
                    "Геология и геохимия нефти и газа (+)(ОПД)",
                    "Геология и геохимия нефти и газа (+)(ОПД)",
                    "Итого за 2 семестр",
                    "Итого за год"
                };

                // Заполнение служебных ячеек: нумерация строк и статические заголовки
                for (int row = 3; row <= 8; row++)
                {
                    // Установка номеров строк в соответствии с логикой нумерации
                    if (row == 3)
                        worksheet.Cells[row, 1] = 1;
                    else if (row == 5)
                        worksheet.Cells[row, 1] = 1;
                    else if (row == 6)
                        worksheet.Cells[row, 1] = 2;
                    else if (row == 7) // Убрано заполнение ячейки A7
                        worksheet.Cells[row, 1] = "";
                    else if (row == 8) // Убрано заполнение ячейки A8
                        worksheet.Cells[row, 1] = "";
                    else
                        worksheet.Cells[row, 1] = row - 2;

                    // Заполнение заголовков строк из массива defaultRowTitles
                    if (row - 3 < defaultRowTitles.Length)
                        worksheet.Cells[row, 2] = defaultRowTitles[row - 3];

                    // Дополнительные описания для конкретных строк с указанием специальностей
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

        // Метод форматирования таблицы Excel для улучшения читаемости и внешнего вида
        private void FormatExcelTableExactly(Ex.Worksheet worksheet)
        {
            try
            {
                // 1. Форматирование первой строки заголовков (общих разделов)
                Ex.Range header1 = worksheet.Range["A1:Z1"];
                header1.Font.Bold = false;
                header1.Font.Size = 11;
                header1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                header1.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                header1.RowHeight = 25;

                // 2. Форматирование второй строки (заголовки столбцов с вертикальным текстом)
                Ex.Range header2 = worksheet.Range["A2:Z2"];
                header2.Orientation = 90; // Поворот текста на 90 градусов для вертикального расположения
                header2.Font.Bold = false;
                header2.Font.Size = 9;
                header2.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;

                header2.WrapText = true; // Включение переноса текста внутри ячейки
                header2.RowHeight = 100; // Увеличенная высота строки для размещения вертикального текста
                header2.Interior.Color = ColorTranslator.ToOle(Color.White); // Установка белого фона

                // 3. Объединение ячеек и специальное форматирование для группировки заголовков
                try
                {
                    // Объединение ячейки "ВСЕГО" (W1:W2) для отображения общего итога - ВЫРАВНИВАНИЕ ВНИЗ
                    Ex.Range rangeW1W2 = worksheet.Range["W1:W2"];
                    rangeW1W2.Merge();
                    rangeW1W2.Orientation = 90;
                    rangeW1W2.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                    rangeW1W2.VerticalAlignment = Ex.XlVAlign.xlVAlignBottom; // Изменено на выравнивание вниз
                    rangeW1W2.Font.Size = 9;
                    rangeW1W2.WrapText = true;

                    // Объединение ячеек A-E по вертикали (заголовки первых пяти столбцов)
                    for (int col = 1; col <= 5; col++)
                    {
                        Ex.Range cellRange = worksheet.Range[worksheet.Cells[1, col], worksheet.Cells[2, col]];
                        cellRange.Merge();
                        cellRange.Orientation = 0; // Горизонтальная ориентация текста
                        cellRange.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                        cellRange.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                        cellRange.Font.Size = 9;
                        cellRange.WrapText = true;

                        // Очистка дублирующихся значений после объединения
                        if (col == 1)
                            worksheet.Cells[1, col] = "N";
                        else
                            worksheet.Cells[1, col] = "";
                        worksheet.Cells[2, col] = "";
                    }

                    // Установка правильных заголовков для объединенных ячеек (основные параметры)
                    worksheet.Cells[1, 2] = "Наименование дисциплины";
                    worksheet.Cells[1, 3] = "Принимаемая нагрузка";
                    worksheet.Cells[1, 4] = "Группы";
                    worksheet.Cells[1, 5] = "Семестр";

                    // Объединение ячеек для заголовка "количество" (F1:H1) - студенты/потоки/группы
                    Ex.Range rangeF1H1 = worksheet.Range["F1:H1"];
                    rangeF1H1.Merge();
                    rangeF1H1.Orientation = 0;
                    rangeF1H1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;

                    // Объединение ячеек для заголовка нагрузки (I1:V1) - виды учебной нагрузки
                    Ex.Range rangeI1V1 = worksheet.Range["I1:V1"];
                    rangeI1V1.Merge();
                    rangeI1V1.Orientation = 0;
                    rangeI1V1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;

                    // Объединение ячеек для заголовка "в том числе" (X1:Z1) - формы обучения
                    Ex.Range rangeX1Z1 = worksheet.Range["X1:Z1"];
                    rangeX1Z1.Merge();
                    rangeX1Z1.Orientation = 0;
                    rangeX1Z1.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при объединении: {ex.Message}");
                }

                // 4. Выделение цветом определенных строк (темно-зеленый с белым текстом) и ЖИРНЫЙ ШРИФТ
                Color darkGreen = Color.FromArgb(0, 100, 0); // Темно-зеленый цвет для выделения
                int darkGreenOle = ColorTranslator.ToOle(darkGreen);

                // Строка 4 (A4:Z4) - "Итого за 1 семестр"
                Ex.Range row4 = worksheet.Range["A4:Z4"];
                row4.Interior.Color = darkGreenOle;
                row4.Font.Color = ColorTranslator.ToOle(Color.White);
                row4.Font.Bold = true; // Добавлен жирный шрифт

                // Строка 7 (A7:Z7) - "Итого за 2 семестр"
                Ex.Range row7 = worksheet.Range["A7:Z7"];
                row7.Interior.Color = darkGreenOle;
                row7.Font.Color = ColorTranslator.ToOle(Color.White);
                row7.Font.Bold = true; // Добавлен жирный шрифт

                // Строка 8 (A8:Z8) - "Итого за год"
                Ex.Range row8 = worksheet.Range["A8:Z8"];
                row8.Interior.Color = darkGreenOle;
                row8.Font.Color = ColorTranslator.ToOle(Color.White);
                row8.Font.Bold = true; // Добавлен жирный шрифт

                // 5. Объединение ячеек для строк с итогами и установка выравнивания влево
                try
                {
                    // Объединяем ячейку B4 с A4 и C4 (A4:C4)
                    Ex.Range rangeB4Merge = worksheet.Range["A4:C4"];
                    // Сохраняем значение из ячейки B4
                    string valueB4 = "";
                    object cellValueB4 = ((Ex.Range)worksheet.Cells[4, 2]).Value2;
                    if (cellValueB4 != null)
                        valueB4 = cellValueB4.ToString();

                    // Очищаем ячейки A4 и C4 - используем прямое присваивание
                    ((Ex.Range)worksheet.Cells[4, 1]).Value2 = "";
                    ((Ex.Range)worksheet.Cells[4, 3]).Value2 = "";

                    // Объединяем ячейки
                    rangeB4Merge.Merge();
                    // Устанавливаем сохраненное значение из B4 в объединенную ячейку
                    rangeB4Merge.Value2 = valueB4;
                    // Выравнивание влево
                    rangeB4Merge.HorizontalAlignment = Ex.XlHAlign.xlHAlignLeft;
                    rangeB4Merge.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                    rangeB4Merge.Font.Bold = true; // Жирный шрифт для объединенной ячейки

                    // Объединяем ячейку B7 с A7 и C7 (A7:C7)
                    Ex.Range rangeB7Merge = worksheet.Range["A7:C7"];
                    // Сохраняем значение из ячейки B7
                    string valueB7 = "";
                    object cellValueB7 = ((Ex.Range)worksheet.Cells[7, 2]).Value2;
                    if (cellValueB7 != null)
                        valueB7 = cellValueB7.ToString();

                    // Очищаем ячейки A7 и C7 - используем прямое присваивание
                    ((Ex.Range)worksheet.Cells[7, 1]).Value2 = "";
                    ((Ex.Range)worksheet.Cells[7, 3]).Value2 = "";

                    // Объединяем ячейки
                    rangeB7Merge.Merge();
                    // Устанавливаем сохраненное значение из B7 в объединенную ячейку
                    rangeB7Merge.Value2 = valueB7;
                    // Выравнивание влево
                    rangeB7Merge.HorizontalAlignment = Ex.XlHAlign.xlHAlignLeft;
                    rangeB7Merge.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                    rangeB7Merge.Font.Bold = true; // Жирный шрифт для объединенной ячейки

                    // Объединяем ячейку B8 с A8 и C8 (A8:C8)
                    Ex.Range rangeB8Merge = worksheet.Range["A8:C8"];
                    // Сохраняем значение из ячейки B8
                    string valueB8 = "";
                    object cellValueB8 = ((Ex.Range)worksheet.Cells[8, 2]).Value2;
                    if (cellValueB8 != null)
                        valueB8 = cellValueB8.ToString();

                    // Очищаем ячейки A8 и C8 - используем прямое присваивание
                    ((Ex.Range)worksheet.Cells[8, 1]).Value2 = "";
                    ((Ex.Range)worksheet.Cells[8, 3]).Value2 = "";

                    // Объединяем ячейки    
                    rangeB8Merge.Merge();
                    // Устанавливаем сохраненное значение из B8 в объединенную ячейку
                    rangeB8Merge.Value2 = valueB8;
                    // Выравнивание влево
                    rangeB8Merge.HorizontalAlignment = Ex.XlHAlign.xlHAlignLeft;
                    rangeB8Merge.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                    rangeB8Merge.Font.Bold = true; // Жирный шрифт для объединенной ячейки
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при объединении итоговых строк: {ex.Message}");
                }

                // 6. Настройка числового формата и выравнивания для всех столбцов
                // ИСКЛЮЧАЕМ строки 4, 7, 8 для столбцов A-C, чтобы не перезаписать выравнивание
                for (int col = 1; col <= 26; col++)
                {
                    // Для столбцов A, B, C работаем только со строками 3, 5, 6
                    if (col == 1 || col == 2 || col == 3)
                    {
                        // Создаем диапазон только для строк, которые не являются итоговыми
                        List<Ex.Range> rangesToFormat = new List<Ex.Range>();

                        // Строка 3
                        rangesToFormat.Add(worksheet.Cells[3, col] as Ex.Range);
                        // Строка 5
                        rangesToFormat.Add(worksheet.Cells[5, col] as Ex.Range);
                        // Строка 6
                        rangesToFormat.Add(worksheet.Cells[6, col] as Ex.Range);

                        foreach (Ex.Range cell in rangesToFormat)
                        {
                            if (cell != null)
                            {
                                if (col == 1) // Столбец A - выравнивание по центру для числовых значений
                                {
                                    cell.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                                    cell.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                                }
                                else // Столбцы B и C - выравнивание влево
                                {
                                    cell.HorizontalAlignment = Ex.XlHAlign.xlHAlignLeft;
                                    cell.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                                }
                            }
                        }
                    }
                    else if (col >= 4 && col <= 26) // Столбцы D-Z
                    {
                        // Для всех строк (3-8) устанавливаем выравнивание по центру
                        Ex.Range columnRange = worksheet.Range[worksheet.Cells[3, col], worksheet.Cells[8, col]];
                        columnRange.NumberFormat = "0"; // Формат отображения без десятичных знаков
                        columnRange.HorizontalAlignment = Ex.XlHAlign.xlHAlignCenter;
                        columnRange.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;
                    }
                }

                // 7. Добавление границ таблицы для всех ячеек (A1:Z8) - ЗЕЛЕНАЯ ОБВОДКА
                Ex.Range tableRange = worksheet.Range["A1:Z8"];

                // Настраиваем ВНЕШНИЕ границы таблицы - ЗЕЛЕНЫЕ и ТОЛСТЫЕ
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeTop].LineStyle = Ex.XlLineStyle.xlContinuous;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeTop].Weight = Ex.XlBorderWeight.xlThick;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeTop].Color = ColorTranslator.ToOle(Color.Green);

                tableRange.Borders[Ex.XlBordersIndex.xlEdgeBottom].LineStyle = Ex.XlLineStyle.xlContinuous;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeBottom].Weight = Ex.XlBorderWeight.xlThick;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeBottom].Color = ColorTranslator.ToOle(Color.Green);

                tableRange.Borders[Ex.XlBordersIndex.xlEdgeLeft].LineStyle = Ex.XlLineStyle.xlContinuous;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeLeft].Weight = Ex.XlBorderWeight.xlThick;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeLeft].Color = ColorTranslator.ToOle(Color.Green);

                tableRange.Borders[Ex.XlBordersIndex.xlEdgeRight].LineStyle = Ex.XlLineStyle.xlContinuous;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeRight].Weight = Ex.XlBorderWeight.xlThick;
                tableRange.Borders[Ex.XlBordersIndex.xlEdgeRight].Color = ColorTranslator.ToOle(Color.Green);

                // Настраиваем ВНУТРЕННИЕ границы таблицы - ЗЕЛЕНЫЕ и ТОЛСТЫЕ
                tableRange.Borders[Ex.XlBordersIndex.xlInsideVertical].LineStyle = Ex.XlLineStyle.xlContinuous;
                tableRange.Borders[Ex.XlBordersIndex.xlInsideVertical].Weight = Ex.XlBorderWeight.xlThick;
                tableRange.Borders[Ex.XlBordersIndex.xlInsideVertical].Color = ColorTranslator.ToOle(Color.Green);

                tableRange.Borders[Ex.XlBordersIndex.xlInsideHorizontal].LineStyle = Ex.XlLineStyle.xlContinuous;
                tableRange.Borders[Ex.XlBordersIndex.xlInsideHorizontal].Weight = Ex.XlBorderWeight.xlThick;
                tableRange.Borders[Ex.XlBordersIndex.xlInsideHorizontal].Color = ColorTranslator.ToOle(Color.Green);

                // 8. Настройка ширины столбцов в соответствии с их содержанием
                // Применяем AutoFit для столбцов C, D и E, а затем немного увеличиваем ширину для запаса
                for (int col = 1; col <= 26; col++)
                {
                    Ex.Range column = (Ex.Range)worksheet.Columns[col];

                    if (col == 1 || (col >= 6 && col <= 8))
                    {
                        column.ColumnWidth = 6; // Узкие столбцы для цифровых данных
                    }
                    else if (col == 2) // Наименование дисциплины - увеличенная ширина
                    {
                        column.ColumnWidth = 45; // Увеличенная ширина для полного отображения названий
                    }
                    else if (col == 3) // Принимаемая нагрузка - увеличенная ширина
                    {
                        // Увеличиваем ширину для столбца C и включаем перенос текста
                        column.ColumnWidth = 80; // Значительно увеличенная ширина
                        // Устанавливаем перенос текста для всего столбца C
                        Ex.Range columnCRange = worksheet.Range["C3:C8"];
                        columnCRange.WrapText = true;
                    }
                    else if (col == 4) // Группы - увеличенная ширина
                    {
                        column.ColumnWidth = 25; // Увеличенная ширина для групп
                        // Устанавливаем перенос текста для всего столбца D
                        Ex.Range columnDRange = worksheet.Range["D3:D8"];
                        columnDRange.WrapText = true;
                    }
                    else if (col == 5) // Семестр - увеличенная ширина
                    {
                        column.ColumnWidth = 20; // Увеличенная ширина для семестра
                        // Устанавливаем перенос текста для всего столбца E
                        Ex.Range columnERange = worksheet.Range["E3:E8"];
                        columnERange.WrapText = true;
                    }
                    else if (col >= 9 && col <= 22) // Столбцы с часами нагрузки
                    {
                        column.ColumnWidth = 4; // Очень узкие столбцы для цифр часов
                    }
                    else if (col == 23) // Столбец W (ВСЕГО) - увеличенная ширина
                    {
                        column.ColumnWidth = 10; // Увеличенная ширина для итогового значения
                    }
                    else if (col >= 24 && col <= 26) // Столбцы X, Y, Z (формы обучения)
                    {
                        column.ColumnWidth = 6; // Увеличенная ширина для названий форм обучения
                    }
                }

                // 9. Автоматическая подгонка ширины столбцов C, D, E для полного отображения текста
                try
                {
                    // Применяем AutoFit для столбцов C, D, E
                    Ex.Range columnC = (Ex.Range)worksheet.Columns[3];
                    Ex.Range columnD = (Ex.Range)worksheet.Columns[4];
                    Ex.Range columnE = (Ex.Range)worksheet.Columns[5];

                    columnC.AutoFit();
                    columnD.AutoFit();
                    columnE.AutoFit();

                    // Добавляем дополнительное пространство после AutoFit для гарантии
                    // Преобразуем ColumnWidth в double для арифметических операций
                    double widthC = Convert.ToDouble(columnC.ColumnWidth);

                    columnC.ColumnWidth = 70;


                    columnD.ColumnWidth = 10;




                    columnE.ColumnWidth = 10;

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при автонастройке ширины столбцов: {ex.Message}");
                }

                // 10. Настройка высоты строк с данными и выравнивания по вертикали
                // Увеличиваем высоту строк для лучшего отображения многострочного текста
                for (int row = 3; row <= 8; row++)
                {
                    Ex.Range rowRange = (Ex.Range)worksheet.Rows[row];
                    // Увеличиваем высоту строк для отображения перенесенного текста
                    rowRange.RowHeight = 45; // Значительно увеличенная высота для многострочного текста
                    rowRange.VerticalAlignment = Ex.XlVAlign.xlVAlignCenter;

                    // Для строк с объединенными ячейками (4, 7, 8) устанавливаем дополнительное форматирование
                    if (row == 4 || row == 7 || row == 8)
                    {
                        // Устанавливаем перенос текста для объединенных ячеек
                        Ex.Range mergedCell = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row, 3]];
                        mergedCell.WrapText = true;
                    }
                }

                // 11. Настройка шрифта для всей таблицы (единообразие оформления)
                Ex.Range allCells = worksheet.Range["A1:Z8"];
                allCells.Font.Name = "Calibri"; // Современный шрифт для хорошей читаемости
                allCells.Font.Size = 11; // Стандартный размер шрифта

                // 12. Фиксация высоты строк заголовков (первая и вторая строки)
                ((Ex.Range)worksheet.Rows[1]).RowHeight = 25; // Стандартная высота для первой строки
                ((Ex.Range)worksheet.Rows[2]).RowHeight = 100; // Увеличенная высота для вертикального текста

                // 13. Изменение цвета текста для определенных ячеек (зеленый цвет)
                Color greenColor = Color.Green;
                int greenOle = ColorTranslator.ToOle(greenColor);

                // Ячейка B3 ("Дополнительно") - зеленый текст
                Ex.Range cellB3 = worksheet.Cells[3, 2] as Ex.Range;
                if (cellB3 != null)
                {
                    cellB3.Font.Color = greenOle;
                }

                // Ячейка B5 (первая "Геология и геохимия нефти и газа") - зеленый текст
                Ex.Range cellB5 = worksheet.Cells[5, 2] as Ex.Range;
                if (cellB5 != null)
                {
                    cellB5.Font.Color = greenOle;
                }

                // Ячейка B6 (вторая "Геология и геохимия нефти и газа") - зеленый текст
                Ex.Range cellB6 = worksheet.Cells[6, 2] as Ex.Range;
                if (cellB6 != null)
                {
                    cellB6.Font.Color = greenOle;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка форматирования: {ex.Message}");
            }
        }

        // Обработчик кнопки "Ввод данных" - открытие CSV файла для редактирования
        private void Data_entry_Click(object sender, EventArgs e)
        {
            try
            {
                // Если CSV файл не существует, создаем пустой файл с указанной кодировкой
                if (!File.Exists(csvFilePath))
                {
                    File.WriteAllText(csvFilePath, "", Encoding.GetEncoding(1251));
                }

                // Открытие файла в программе по умолчанию (обычно Блокнот или Excel)
                Process.Start(csvFilePath);

                // Вывод инструкций для пользователя по редактированию данных
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

        // Обработчик кнопки "Выход" - завершение работы приложения
        private void exit_Click(object sender, EventArgs e)
        {
            Application.Exit(); // Завершение работы всего приложения
        }

        // Обработчик события загрузки формы (в текущей реализации не содержит кода)
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // Обработчик кнопки "Информация" - отображение справки по использованию программы
        private void info_Click(object sender, EventArgs e)
        {
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
    }
}