
using System;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace WordWork19
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Переменные для хранения текста строк титульника
        public string ministryLine = "Министерство транспорта Российской Федерации";
        public string federalLine = "Федеральное государственное автономное образовательное";
        public string institutionLine = "учреждение высшего образования";
        public string universityLine = "«Российский университет транспорта»";
        public string acronymLine = "(ФГАОУ ВО РУТ(МИИТ), РУТ (МИИТ)";
        public string instituteLine = "Институт транспортной техники и систем управления";
        public string departmentLine = "Кафедра «Управление и защита информации»";
        public string labWorkLine = "Лабораторная работа № 19";
        public string disciplineLine = "по дисциплине: «Программирование и основы алгоритмизации»";
        public string topicLine = "на тему: «программируемая настройка параметров документов *Microsoft Office Word*»";
        public string performedByLine = "Выполнил: ст. гр. ТУУ-111";
        public string studentNameLine = "Михалин А.В";
        public string variantLine = "Вариант №2";
        public string dateLine = "29.05.2025";
        public string checkedByLine = "Проверил: к.т.н., доц. Сафронов А.И.";
        public string cityYearLine = "Москва - 2025 г.";

        // Цвета для каждой строки титульника (как WdColor)
        private Word.WdColor[] lineWordColors = new Word.WdColor[16];

        // Текущая строка для редактирования (0-15)
        private int currentLineIndex = 0;

        // Путь к последнему созданному файлу
        private string lastCreatedFilePath = "";

        // Массив названий строк для отображения
        private string[] lineNames =
        {
            "1. Министерство транспорта Российской Федерации",
            "2. Федеральное государственное автономное образовательное",
            "3. учреждение высшего образования",
            "4. «Российский университет транспорта»",
            "5. (ФГАОУ ВО РУТ(МИИТ), РУТ (МИИТ)",
            "6. Институт транспортной техники и систем управления",
            "7. Кафедра «Управление и защита информации»",
            "8. Лабораторная работа № 19",
            "9. по дисциплине: «Программирование и основы алгоритмизации»",
            "10. на тему: «программируемая настройка параметров документов *Microsoft Office Word*»",
            "11. Выполнил: ст. гр. ТУУ-111",
            "12. Михалин А.В",
            "13. Вариант №2",
            "14. 29.05.2025",
            "15. Проверил: к.т.н., доц. Сафронов А.И.",
            "16. Москва - 2025 г."
        };

        // Текстовые значения по умолчанию
        private string[] defaultLines =
        {
            "Министерство транспорта Российской Федерации",
            "Федеральное государственное автономное образовательное",
            "учреждение высшего образования",
            "«Российский университет транспорта»",
            "(ФГАОУ ВО РУТ(МИИТ), РУТ (МИИТ)",
            "Институт транспортной техники и систем управления",
            "Кафедра «Управление и защита информации»",
            "Лабораторная работа № 19",
            "по дисциплине: «Программирование и основы алгоритмизации»",
            "на тему: «программируемая настройка параметров документов *Microsoft Office Word*»",
            "Выполнил: ст. гр. ТУУ-111",
            "Михалин А.В",
            "Вариант №2",
            "29.05.2025",
            "Проверил: к.т.н., доц. Сафронов А.И.",
            "Москва - 2025 г."
        };

        private void Form1_Load(object sender, EventArgs e)
        {
            // Инициализация цветов (по умолчанию черный)
            for (int i = 0; i < 16; i++)
            {
                lineWordColors[i] = Word.WdColor.wdColorBlack;
            }

            // Настройка ComboBox с цветами
            checkColor1.Items.Add("Черный");
            checkColor1.Items.Add("Красный");
            checkColor1.Items.Add("Синий");
            checkColor1.Items.Add("Зеленый");
            checkColor1.Items.Add("Фиолетовый");
            checkColor1.SelectedIndex = 0;

            // Скрываем элементы ввода до нажатия кнопки FormFile
            checkColor1.Visible = false;
            Input.Visible = false;
            input1.Visible = false;
            richTextBox1.Visible = false;

            // Кнопки Назад и Просмотр по умолчанию невидимы
            Back.Visible = false;
            Preview.Visible = false;

            // Устанавливаем начальный заголовок формы
            this.Text = "Титульник лабораторной работы";
        }

        private void FormFile_Click(object sender, EventArgs e)
        {
            // Показываем элементы ввода
            checkColor1.Visible = true;
            Input.Visible = true;
            input1.Visible = true;
            richTextBox1.Visible = true;

            // Активируем и делаем видимой кнопку Назад
            Back.Visible = true;
            Back.Enabled = true;

            // Начинаем с первой строки
            currentLineIndex = 0;

            // Отображаем текущую строку в richTextBox
            UpdateRichTextBox();

            // Устанавливаем значение в поле ввода
            Input.Text = GetCurrentLineValue();

            // Устанавливаем фокус на поле ввода
            Input.Focus();
        }

        private string GetCurrentLineValue()
        {
            // Получаем значение текущей строки
            switch (currentLineIndex)
            {
                case 0: return ministryLine;
                case 1: return federalLine;
                case 2: return institutionLine;
                case 3: return universityLine;
                case 4: return acronymLine;
                case 5: return instituteLine;
                case 6: return departmentLine;
                case 7: return labWorkLine;
                case 8: return disciplineLine;
                case 9: return topicLine;
                case 10: return performedByLine;
                case 11: return studentNameLine;
                case 12: return variantLine;
                case 13: return dateLine;
                case 14: return checkedByLine;
                case 15: return cityYearLine;
                default: return "";
            }
        }

        private void UpdateRichTextBox()
        {
            richTextBox1.Clear();

            // Отображаем все строки
            for (int i = 0; i < 16; i++)
            {
                string lineValue = GetLineValue(i);

                // Формируем строку
                string line = $"{lineNames[i]}: {lineValue}";

                // Если это текущая редактируемая строка, выделяем ее
                if (i == currentLineIndex)
                {
                    richTextBox1.SelectionColor = System.Drawing.Color.Red;
                    richTextBox1.AppendText($"→ {line}\n");
                    richTextBox1.SelectionColor = richTextBox1.ForeColor;
                }
                else
                {
                    richTextBox1.AppendText($"  {line}\n");
                }
            }

            // Прокручиваем к текущей строке
            richTextBox1.ScrollToCaret();

            // Обновляем заголовок формы с информацией о текущей строке
            // Добавляем проверку на границы массива
            if (currentLineIndex >= 0 && currentLineIndex < lineNames.Length)
            {
                this.Text = $"Титульник лабораторной работы - Текущая строка: {lineNames[currentLineIndex]}";
            }
            else
            {
                this.Text = "Титульник лабораторной работы";
            }

            // Обновляем выбранный цвет для текущей строки
            UpdateSelectedColor();
        }

        private void UpdateSelectedColor()
        {
            // Устанавливаем выбранный цвет в ComboBox на основе текущей строки
            // Добавляем проверку на границы массива
            if (currentLineIndex >= 0 && currentLineIndex < lineWordColors.Length)
            {
                Word.WdColor currentColor = lineWordColors[currentLineIndex];

                int selectedIndex = 0; // По умолчанию черный

                if (currentColor == Word.WdColor.wdColorRed)
                    selectedIndex = 1;
                else if (currentColor == Word.WdColor.wdColorBlue)
                    selectedIndex = 2;
                else if (currentColor == Word.WdColor.wdColorGreen)
                    selectedIndex = 3;
                else if (currentColor == Word.WdColor.wdColorViolet)
                    selectedIndex = 4;

                checkColor1.SelectedIndex = selectedIndex;
            }
        }

        private string GetLineValue(int index)
        {
            // Получаем значение строки по индексу
            switch (index)
            {
                case 0: return ministryLine;
                case 1: return federalLine;
                case 2: return institutionLine;
                case 3: return universityLine;
                case 4: return acronymLine;
                case 5: return instituteLine;
                case 6: return departmentLine;
                case 7: return labWorkLine;
                case 8: return disciplineLine;
                case 9: return topicLine;
                case 10: return performedByLine;
                case 11: return studentNameLine;
                case 12: return variantLine;
                case 13: return dateLine;
                case 14: return checkedByLine;
                case 15: return cityYearLine;
                default: return "";
            }
        }

        private void SetLineValue(int index, string value)
        {
            // Устанавливаем значение строки по индексу
            switch (index)
            {
                case 0: ministryLine = value; break;
                case 1: federalLine = value; break;
                case 2: institutionLine = value; break;
                case 3: universityLine = value; break;
                case 4: acronymLine = value; break;
                case 5: instituteLine = value; break;
                case 6: departmentLine = value; break;
                case 7: labWorkLine = value; break;
                case 8: disciplineLine = value; break;
                case 9: topicLine = value; break;
                case 10: performedByLine = value; break;
                case 11: studentNameLine = value; break;
                case 12: variantLine = value; break;
                case 13: dateLine = value; break;
                case 14: checkedByLine = value; break;
                case 15: cityYearLine = value; break;
            }
        }

        private void checkColor1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Этот метод вызывается при изменении выбора цвета
            // Если пользователь меняет цвет при возврате к строке, обновляем цвет
            if (Back.Enabled && checkColor1.Visible)
            {
                // Получаем выбранный цвет Word
                Word.WdColor selectedWordColor = Word.WdColor.wdColorBlack;
                switch (checkColor1.SelectedItem.ToString())
                {
                    case "Черный": selectedWordColor = Word.WdColor.wdColorBlack; break;
                    case "Красный": selectedWordColor = Word.WdColor.wdColorRed; break;
                    case "Синий": selectedWordColor = Word.WdColor.wdColorBlue; break;
                    case "Зеленый": selectedWordColor = Word.WdColor.wdColorGreen; break;
                    case "Фиолетовый": selectedWordColor = Word.WdColor.wdColorViolet; break;
                }

                // Сохраняем новый цвет для текущей строки (с проверкой границ)
                if (currentLineIndex >= 0 && currentLineIndex < lineWordColors.Length)
                {
                    lineWordColors[currentLineIndex] = selectedWordColor;
                }

                // Обновляем отображение в richTextBox
                UpdateRichTextBox();
            }
        }

        private void Input_TextChanged(object sender, EventArgs e)
        {
            // Этот метод вызывается при изменении текста в поле ввода
        }

        private void input1_Click(object sender, EventArgs e)
        {
            // Сохраняем введенный текст в текущую строку
            string inputText = Input.Text.Trim();

            if (string.IsNullOrEmpty(inputText))
            {
                MessageBox.Show("Введите текст для строки титульника!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Получаем выбранный цвет Word
            Word.WdColor selectedWordColor = Word.WdColor.wdColorBlack;
            switch (checkColor1.SelectedItem.ToString())
            {
                case "Черный": selectedWordColor = Word.WdColor.wdColorBlack; break;
                case "Красный": selectedWordColor = Word.WdColor.wdColorRed; break;
                case "Синий": selectedWordColor = Word.WdColor.wdColorBlue; break;
                case "Зеленый": selectedWordColor = Word.WdColor.wdColorGreen; break;
                case "Фиолетовый": selectedWordColor = Word.WdColor.wdColorViolet; break;
            }

            // Сохраняем значение строки (с проверкой границ)
            if (currentLineIndex >= 0 && currentLineIndex < 16)
            {
                SetLineValue(currentLineIndex, inputText);
                lineWordColors[currentLineIndex] = selectedWordColor;
            }

            // Переходим к следующей строке
            currentLineIndex++;

            // Если все строки заполнены
            if (currentLineIndex >= 16)
            {
                MessageBox.Show("Все строки титульника заполнены! Теперь можно создать документ Word.",
                    "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Скрываем элементы ввода
                checkColor1.Visible = false;
                Input.Visible = false;
                input1.Visible = false;

                // Обновляем отображение (устанавливаем заголовок без текущей строки)
                UpdateRichTextBoxWithoutCurrentLine();
                return;
            }

            // Устанавливаем значение следующей строки в поле ввода
            Input.Text = GetCurrentLineValue();

            // Обновляем отображение
            UpdateRichTextBox();

            // Устанавливаем фокус на поле ввода
            Input.Focus();
        }

        private void UpdateRichTextBoxWithoutCurrentLine()
        {
            richTextBox1.Clear();

            // Отображаем все строки
            for (int i = 0; i < 16; i++)
            {
                string lineValue = GetLineValue(i);
                string line = $"{lineNames[i]}: {lineValue}";
                richTextBox1.AppendText($"  {line}\n");
            }

            // Прокручиваем к началу
            richTextBox1.ScrollToCaret();

            // Устанавливаем общий заголовок
            this.Text = "Титульник лабораторной работы - Все строки заполнены";
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            // Этот метод вызывается при изменении текста в richTextBox
        }

        private void CreateWord_Click(object sender, EventArgs e)
        {
            Word.Application oWord = null;
            Word.Document oDoc = null;

            try
            {
                oWord = new Word.Application();
                oWord.Visible = false;
                oDoc = oWord.Documents.Add();

                // Устанавливаем шрифт Times New Roman для всего документа
                oDoc.Content.Font.Name = "Times New Roman";
                oDoc.Content.Font.Size = 14;

                // 1. Министерство транспорта Российской Федерации - по центру
                Word.Paragraph line1 = oDoc.Paragraphs.Add();
                line1.Range.Font.Size = 14;
                line1.Range.Text = ministryLine;
                line1.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line1.Range.InsertParagraphAfter();

                // 2. Федеральное государственное автономное образовательное - по центру
                Word.Paragraph line2 = oDoc.Paragraphs.Add();
                line2.Range.Font.Size = 14;
                line2.Range.Text = federalLine;
                line2.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line2.Range.InsertParagraphAfter();

                // 3. учреждение высшего образования - по центру
                Word.Paragraph line3 = oDoc.Paragraphs.Add();
                line3.Range.Font.Size = 14;
                line3.Range.Text = institutionLine;
                line3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line3.Range.InsertParagraphAfter();

                // 4. «Российский университет транспорта» - по центру
                Word.Paragraph line4 = oDoc.Paragraphs.Add();
                line4.Range.Font.Size = 14;
                line4.Range.Text = universityLine;
                line4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line4.Range.InsertParagraphAfter();
                // 5. (ФГАОУ ВО РУТ(МИИТ), РУТ (МИИТ) - по центру с подчеркиванием во всю ширину страницы

                // Создаем таблицу с 1 строкой и 1 ячейкой
                object missing = System.Reflection.Missing.Value;
                object autoFitBehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;

                Word.Table underlineTable = oDoc.Tables.Add(
                    oDoc.Paragraphs.Add().Range,
                    1, // строки
                    1, // столбцы
                    ref missing,
                    ref missing);

                // Устанавливаем текст
                underlineTable.Cell(1, 1).Range.Text = acronymLine;
                underlineTable.Cell(1, 1).Range.Font.Size = 14;
                underlineTable.Cell(1, 1).Range.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Отключаем все границы таблицы
                underlineTable.Borders.Enable = 0;

                // Включаем только нижнюю границу с явным приведением типов
                Word.Border bottomBorder = underlineTable.Borders[(Microsoft.Office.Interop.Word.WdBorderType)Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom];
                bottomBorder.LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                bottomBorder.LineWidth = Microsoft.Office.Interop.Word.WdLineWidth.wdLineWidth050pt;
                bottomBorder.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;

                // Отключаем остальные границы таблицы
                underlineTable.Borders[(Microsoft.Office.Interop.Word.WdBorderType)Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop].LineStyle =
                    Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone;
                underlineTable.Borders[(Microsoft.Office.Interop.Word.WdBorderType)Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft].LineStyle =
                    Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone;
                underlineTable.Borders[(Microsoft.Office.Interop.Word.WdBorderType)Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight].LineStyle =
                    Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone;

                // Убираем отступы
                underlineTable.LeftPadding = 0;
                underlineTable.RightPadding = 0;
                underlineTable.TopPadding = 0;
                underlineTable.BottomPadding = 0;

                // Автоматическая подгонка ширины
                underlineTable.AutoFitBehavior((Microsoft.Office.Interop.Word.WdAutoFitBehavior)autoFitBehavior);

                // Переходим к следующему абзацу
                oDoc.Paragraphs.Add();

                // 6. Институт транспортной техники и систем управления - по центру
                Word.Paragraph line6 = oDoc.Paragraphs.Add();
                line6.Range.Font.Size = 14;
                line6.Range.Text = instituteLine;
                line6.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line6.Range.InsertParagraphAfter();





                // 7. Кафедра «Управление и защита информации» - по центру
                Word.Paragraph line7 = oDoc.Paragraphs.Add();
                line7.Range.Font.Size = 14;
                line7.Range.Text = departmentLine;
                line7.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line7.Range.InsertParagraphAfter();

                // Пустая строка
                oDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                // 8. Лабораторная работа № 19 - по центру, жирный, размер 28
                Word.Paragraph line8 = oDoc.Paragraphs.Add();

                line8.Range.Font.Size = 28; // Изменено с 16 на 28
                line8.Range.Text = labWorkLine;
                line8.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line8.Range.InsertParagraphAfter();

                // Пустая строка
                oDoc.Paragraphs.Add().Range.InsertParagraphAfter();

                // 9. по дисциплине: «Программирование и основы алгоритмизации» - по центру
                Word.Paragraph line9 = oDoc.Paragraphs.Add();
                line9.Range.Font.Size = 14;
                line9.Range.Text = disciplineLine;
                line9.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line9.Range.InsertParagraphAfter();

                // 10. на тему: «программируемая настройка параметров документов *Microsoft Office Word*» - по центру
                Word.Paragraph line10 = oDoc.Paragraphs.Add();
                line10.Range.Font.Size = 14;
                line10.Range.Text = topicLine;
                line10.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line10.Range.InsertParagraphAfter();

                // Несколько пустых строк
                for (int i = 0; i < 1; i++)
                {
                    oDoc.Paragraphs.Add().Range.InsertParagraphAfter();
                }

                // 11. Выполнил: ст. гр. ТУУ-111 - выравнивание вправо
                Word.Paragraph line11 = oDoc.Paragraphs.Add();
                line11.Range.Font.Size = 14;
                line11.Range.Text = performedByLine;
                line11.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                line11.Range.InsertParagraphAfter();

                // 12. Михалин А.В - выравнивание вправо
                Word.Paragraph line12 = oDoc.Paragraphs.Add();
                line12.Range.Font.Size = 14;
                line12.Range.Text = studentNameLine;
                line12.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                line12.Range.InsertParagraphAfter();

                // 13. Вариант №2 - выравнивание вправо
                Word.Paragraph line13 = oDoc.Paragraphs.Add();
                line13.Range.Font.Size = 14;
                line13.Range.Text = variantLine;
                line13.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                line13.Range.InsertParagraphAfter();

                // 14. 29.05.2025 - выравнивание вправо
                Word.Paragraph line14 = oDoc.Paragraphs.Add();
                line14.Range.Font.Size = 14;
                line14.Range.Text = dateLine;
                line14.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                line14.Range.InsertParagraphAfter();

                // Добавляем надстрочную надпись "(дата выполнения)" после даты
                Word.Paragraph dateExecutionSuperscript = oDoc.Paragraphs.Add();
                dateExecutionSuperscript.Range.Font.Size = 14; // Меньший размер для надстрочной надписи
                dateExecutionSuperscript.Range.Font.Superscript = 1; // Включаем надстрочное начертание
                dateExecutionSuperscript.Range.Text = "(дата выполнения)";
                dateExecutionSuperscript.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                dateExecutionSuperscript.Range.InsertParagraphAfter();

                // 15. Проверил: к.т.н., доц. Сафронов А.И. - выравнивание вправо
                // Без пустых строк между 14 и 15 строками
                Word.Paragraph line15 = oDoc.Paragraphs.Add();
                line15.Range.Font.Size = 14;
                line15.Range.Text = checkedByLine;
                line15.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                line15.Range.InsertParagraphAfter();

                // Добавляем надстрочную надпись "(дата приёмки)" после проверяющего
                Word.Paragraph dateAcceptanceSuperscript = oDoc.Paragraphs.Add();
                dateAcceptanceSuperscript.Range.Font.Size = 14; // Меньший размер для надстрочной надписи
                dateAcceptanceSuperscript.Range.Font.Superscript = 1; // Включаем надстрочное начертание
                dateAcceptanceSuperscript.Range.Text = "(дата приёмки)";
                dateAcceptanceSuperscript.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                dateAcceptanceSuperscript.Range.InsertParagraphAfter();

                // Несколько пустых строк перед городом и годом
                for (int i = 0; i < 1; i++)
                {
                    oDoc.Paragraphs.Add().Range.InsertParagraphAfter();
                }


                // 16. Москва -- 2025 г. - по центру
                Word.Paragraph line16 = oDoc.Paragraphs.Add();
                line16.Range.Font.Size = 14;
                line16.Range.Text = cityYearLine;
                line16.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                line16.Range.InsertParagraphAfter();

                // Применяем цвета к строкам
                ApplyColorsToLines(oDoc);

                // Сохраняем документ с уникальным именем (временная метка)
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"Титульник_Лабораторной_19_{timestamp}.docx";
                string filePath = Path.Combine(System.Windows.Forms.Application.StartupPath, fileName);

                // Сохраняем путь к последнему созданному файлу
                lastCreatedFilePath = filePath;

                // Удаляем файл если он уже существует
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }

                // Используем SaveAs 
                oDoc.SaveAs(filePath, Word.WdSaveFormat.wdFormatDocumentDefault);

                // Закрываем документ
                oDoc.Close(false);
                oWord.Quit();

                // Активируем и делаем видимой кнопку Просмотр
                Preview.Visible = true;
                Preview.Enabled = true;

                MessageBox.Show($"Титульник лабораторной работы создан: {filePath}\n\nКаждое нажатие кнопки создает новый файл с уникальным именем.",
                    "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании документа: {ex.Message}\n\nПодробности: {ex.ToString()}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // В случае ошибки пытаемся закрыть Word
                try
                {
                    if (oWord != null)
                    {
                        oWord.Quit();
                    }
                }
                catch { }
            }
            finally
            {
                // Сборка мусора для освобождения COM-объектов
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }

        // Метод для применения цветов к строкам
        private void ApplyColorsToLines(Word.Document doc)
        {
            try
            {
                // Массив значений строк
                string[] lineValues = {
                    ministryLine, federalLine, institutionLine, universityLine,
                    acronymLine, instituteLine, departmentLine, labWorkLine,
                    disciplineLine, topicLine, performedByLine, studentNameLine,
                    variantLine, dateLine, checkedByLine, cityYearLine
                };

                // Ищем каждую строку в документе и применяем цвет
                for (int i = 0; i < 16; i++)
                {
                    if (!string.IsNullOrEmpty(lineValues[i]))
                    {
                        Word.Range searchRange = doc.Content;
                        searchRange.Find.ClearFormatting();
                        searchRange.Find.Text = lineValues[i];

                        while (searchRange.Find.Execute())
                        {
                            // Применяем цвет к найденному тексту
                            searchRange.Font.Color = lineWordColors[i];

                            // Ищем следующее вхождение
                            searchRange = doc.Range(searchRange.End, doc.Content.End);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Игнорируем ошибки при применении цветов
                Console.WriteLine($"Ошибка при применении цветов: {ex.Message}");
            }
        }

        private void ResetButton_Click(object sender, EventArgs e)
        {
            // Сброс всех значений к исходным
            for (int i = 0; i < 16; i++)
            {
                SetLineValue(i, defaultLines[i]);
                lineWordColors[i] = Word.WdColor.wdColorBlack;
            }

            // Сброс текущего индекса
            currentLineIndex = 0;

            // Обновление отображения
            UpdateRichTextBox();

            // Скрытие элементов ввода
            checkColor1.Visible = false;
            Input.Visible = false;
            input1.Visible = false;
            richTextBox1.Visible = false;

            // Скрытие кнопок
            Back.Visible = false;
            Preview.Visible = false;

            // Сброс пути к последнему файлу
            lastCreatedFilePath = "";

            // Восстановление заголовка формы
            this.Text = "Титульник лабораторной работы";

            MessageBox.Show("Все значения сброшены к исходным. Нажмите 'Ввод строк титульника' для начала редактирования.",
                "Сброс", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Back_Click(object sender, EventArgs e)
        {
            // Проверяем, не выходим ли мы за границы
            if (currentLineIndex > 0)
            {
                // Возвращаемся на предыдущую строку
                currentLineIndex--;

                // Показываем элементы ввода, если они были скрыты
                checkColor1.Visible = true;
                Input.Visible = true;
                input1.Visible = true;
                richTextBox1.Visible = true;

                // Устанавливаем значение текущей строки в поле ввода
                Input.Text = GetCurrentLineValue();

                // Обновляем отображение
                UpdateRichTextBox();

                // Устанавливаем фокус на поле ввода
                Input.Focus();
            }
            else
            {
                MessageBox.Show("Вы находитесь на первой строке. Возврат невозможен.",
                    "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Preview_Click(object sender, EventArgs e)
        {
            // Проверяем, существует ли последний созданный файл
            if (string.IsNullOrEmpty(lastCreatedFilePath) || !File.Exists(lastCreatedFilePath))
            {
                MessageBox.Show("Файл для предпросмотра не найден. Сначала создайте документ Word.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Открываем файл в ассоциированной программе (Microsoft Word)
                Process.Start(lastCreatedFilePath);

                MessageBox.Show($"Файл открыт для предпросмотра: {lastCreatedFilePath}",
                    "Предпросмотр", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии файла: {ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       
        private void CreateReceipt_Click(object sender, EventArgs e)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;
            bool fileCopied = false;

            try
            {
                // Проверяем наличие файла-шаблона
                string sourceFile = "Квитанцияобр167.docx";
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, sourceFile);

                // Создаем имя для нового файла
                string newFileName = $"Квитанция_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                string newFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, newFileName);

                // Создаем приложение Word
                wordApp = new Word.Application();
                wordApp.Visible = true; // Показываем для отладки

                if (File.Exists(templatePath))
                {
                    
                    File.Copy(templatePath, newFilePath, overwrite: true);
                    fileCopied = true;

                    // Открываем новый файл
                    doc = wordApp.Documents.Open(newFilePath);

                   
                }
                else
                {
                    doc = wordApp.Documents.Add();

                    // Устанавливаем размер страницы А5
                    doc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                    doc.PageSetup.PageHeight = 595.35f;  // A5 высота 21 см
                    doc.PageSetup.PageWidth = 419.58f;   // A5 ширина 14.8 см

                    // Настраиваем поля
                    doc.PageSetup.TopMargin = 42.55f;    // 1.5 см в пунктах
                    doc.PageSetup.BottomMargin = 42.55f;
                    doc.PageSetup.LeftMargin = 42.55f;
                    doc.PageSetup.RightMargin = 42.55f;

                    // =========== СОЗДАЕМ КВИТАНЦИЮ БЕЗ ТАБЛИЦЫ ===========
                    // Сначала создадим квитанцию без таблицы, чтобы убедиться, что все работает

                    // Заголовок
                    Word.Paragraph titlePara = doc.Paragraphs.Add();
                    titlePara.Range.Text = "КВИТАНЦИЯ №";
                    titlePara.Range.Font.Bold = 1;
                    titlePara.Range.Font.Size = 12;
                    titlePara.Range.InsertParagraphAfter();

                    // Дата
                    Word.Paragraph datePara = doc.Paragraphs.Add();
                    datePara.Range.Text = "«____» ______ 20 ____ г.";
                    datePara.Range.Font.Size = 12;
                    datePara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    datePara.Range.InsertParagraphAfter();

                    // Пустая строка
                    doc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Серия АС
                    Word.Paragraph seriesPara = doc.Paragraphs.Add();
                    seriesPara.Range.Text = "Серия АС";
                    seriesPara.Range.Font.Size = 12;
                    seriesPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    seriesPara.Range.InsertParagraphAfter();

                    // Пустая строка
                    doc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Учреждение
                    Word.Paragraph institutionPara = doc.Paragraphs.Add();
                    institutionPara.Range.Text = "Учреждение МБДОУ ДЕТСКИЙ САД «КОЛОКОЛЬЧИК»";
                    // Подчеркиваем только название
                    Word.Range institutionRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                    institutionRange.Text = "МБДОУ ДЕТСКИЙ САД «КОЛОКОЛЬЧИК»";
                    institutionRange.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    institutionRange.Font.Size = 12;
                    institutionRange.InsertParagraphAfter();

                    // Разделительная линия
                    Word.Paragraph linePara = doc.Paragraphs.Add();
                    linePara.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    linePara.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                    linePara.Range.Text = "";
                    linePara.Range.InsertParagraphAfter();

                    // Местонахождение
                    Word.Paragraph locationPara = doc.Paragraphs.Add();
                    locationPara.Range.Text = "Местонахождение";
                    locationPara.Range.Font.Bold = 1;
                    locationPara.Range.Font.Size = 12;
                    locationPara.Range.InsertParagraphAfter();

                    // Адрес
                    Word.Paragraph addressPara = doc.Paragraphs.Add();
                    addressPara.Range.Text = "Тербунский район, с. Тербуны, ул. Коммунальная, д. 15";
                    addressPara.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    addressPara.Range.Font.Size = 12;
                    addressPara.Range.InsertParagraphAfter();

                    // Пустая строка
                    doc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Принято от
                    Word.Paragraph acceptedPara = doc.Paragraphs.Add();
                    acceptedPara.Range.Text = "Принято от ________________________";
                    acceptedPara.Range.Font.Size = 12;
                    acceptedPara.Range.InsertParagraphAfter();

                    Word.Paragraph acceptedDesc = doc.Paragraphs.Add();
                    acceptedDesc.Range.Text = "(фамилия, имя, отчество)";
                    acceptedDesc.Range.Font.Size = 10;
                    acceptedDesc.Range.InsertParagraphAfter();

                    // В уплату
                    Word.Paragraph paymentPara = doc.Paragraphs.Add();
                    paymentPara.Range.Text = "В уплату родительской платы ________________________";
                    paymentPara.Range.Font.Size = 12;
                    paymentPara.Range.InsertParagraphAfter();

                    Word.Paragraph paymentDesc = doc.Paragraphs.Add();
                    paymentDesc.Range.Text = "(вид продукции, услуги, работы)";
                    paymentDesc.Range.Font.Size = 10;
                    paymentDesc.Range.InsertParagraphAfter();

                    // Сумма
                    Word.Paragraph amountPara = doc.Paragraphs.Add();
                    amountPara.Range.Text = "Сумма, всего ________________________";
                    amountPara.Range.Font.Size = 12;
                    amountPara.Range.InsertParagraphAfter();

                    Word.Paragraph amountDesc = doc.Paragraphs.Add();
                    amountDesc.Range.Text = "(прописью)";
                    amountDesc.Range.Font.Size = 10;
                    amountDesc.Range.InsertParagraphAfter();

                    // Рубли и копейки
                    Word.Paragraph rublesPara = doc.Paragraphs.Add();
                    rublesPara.Range.Text = "________________________ руб. ________________________ коп.";
                    rublesPara.Range.Font.Size = 12;
                    rublesPara.Range.InsertParagraphAfter();

                    // Пустая строка
                    doc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Наличными
                    Word.Paragraph cashPara = doc.Paragraphs.Add();
                    cashPara.Range.Text = "в том числе: наличными деньгами ________________________ (прописью) руб. ________________________ коп.";
                    cashPara.Range.Font.Size = 12;
                    cashPara.Range.InsertParagraphAfter();

                    // Картой
                    Word.Paragraph cardPara = doc.Paragraphs.Add();
                    cardPara.Range.Text = "с использованием платёжной карты ________________________ (прописью) руб. ________________________ коп.";
                    cardPara.Range.Font.Size = 12;
                    cardPara.Range.InsertParagraphAfter();

                    // Пустая строка
                    doc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Получил
                    Word.Paragraph receivedPara = doc.Paragraphs.Add();
                    receivedPara.Range.Text = "Получил ________________________ ________________________ ________________________";
                    receivedPara.Range.Font.Size = 12;
                    receivedPara.Range.InsertParagraphAfter();

                    Word.Paragraph receivedDesc = doc.Paragraphs.Add();
                    receivedDesc.Range.Text = "(должность) (подпись) (расшифровка подписи)";
                    receivedDesc.Range.Font.Size = 10;
                    receivedDesc.Range.InsertParagraphAfter();

                    // Уплатил
                    Word.Paragraph paidPara = doc.Paragraphs.Add();
                    paidPara.Range.Text = "Уплатил ________________________ «____» ______ 20 ____ г.";
                    paidPara.Range.Font.Size = 12;
                    paidPara.Range.InsertParagraphAfter();

                    Word.Paragraph paidDesc = doc.Paragraphs.Add();
                    paidDesc.Range.Text = "(подпись)";
                    paidDesc.Range.Font.Size = 10;
                    paidDesc.Range.InsertParagraphAfter();

                    // Пустая строка
                    doc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // М.П.
                    Word.Paragraph mpPara = doc.Paragraphs.Add();
                    mpPara.Range.Text = "М.П.";
                    mpPara.Range.Font.Bold = 1;
                    mpPara.Range.Font.Size = 12;
                    mpPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    mpPara.Range.InsertParagraphAfter();

                    // Пустая строка
                    doc.Paragraphs.Add().Range.InsertParagraphAfter();

                    // Реквизиты ИП
                    Word.Paragraph footerPara = doc.Paragraphs.Add();
                    footerPara.Range.Text = "ИП Антипина С.Е. ул. А. Гайгеровой, 42, ИНН 482103982900, Заказ № 1816, Формат А5, Бумага офсетная 80 гр./кв. м. Тираж 40×50 шт. 2016 г.";
                    footerPara.Range.Font.Size = 8;
                    footerPara.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    footerPara.Range.InsertParagraphAfter();

                    // Сохраняем новый документ только если он создавался с нуля
                    doc.SaveAs2(newFilePath, Word.WdSaveFormat.wdFormatDocumentDefault);
                }

                // Сообщаем о результате без блокировки потока
                // Не закрываем Word после MessageBox
                MessageBox.Show($"Квитанция успешно создана:\n{newFilePath}",
                    "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Word остается открытым - НЕ освобождаем ресурсы здесь
                return; // Выходим, не закрывая Word
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при создании квитанции:\n{ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                // В случае ошибки закрываем Word
                try
                {
                    if (doc != null)
                    {
                        doc.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                    }
                }
                catch { }

                try
                {
                    if (wordApp != null)
                    {
                        wordApp.Quit(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    }
                }
                catch { }
            }
            finally
            {
                // Принудительная сборка мусора только при ошибке
                // Если все успешно, Word остается открытым
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
    }
}
    