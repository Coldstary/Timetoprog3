using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CsvWork21new
{
    public partial class CSVfilter : Form
    {
        private DataManager dataManager; // Объявляем переменную

        public CSVfilter()
        {
            InitializeComponent();
        }

        private void CSVfilter_Load(object sender, EventArgs e)
        {
            // Инициализируем dataManager при загрузке формы
            try
            {
                dataManager = new DataManager();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Настройка элементов при загрузке
            CmbFilterType.Items.Clear();
            CmbFilterType.Items.AddRange(new object[]
            {
                "Все лекарства",
                "Поиск лекарств (название, форма, производитель)",
                "Лекарства для лечения болезни",
                "Продажи за период",
                "Бронирование лекарства",
                "Поступление со склада"
            });
            CmbFilterType.SelectedIndex = 0;

            // Настройка календаря
            monthCalendar.MaxSelectionCount = 31;
            monthCalendar.Visible = false;

            // Показать все лекарства при загрузке
            ShowAllMedicines();
        }


        private void CmbFilterType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dataManager == null) return;

            // Скрываем все дополнительные элементы
            monthCalendar.Visible = false;
           
            BtnReserve.Visible = false;
            BtnExecute.Visible = true;

            // Сбрасываем текстовые поля
            TxtSearchParam1.Clear();
             TxtSearchParam2.Clear();
            TxtSearchParam3.Clear();

            switch (CmbFilterType.SelectedIndex)
            {
                case 0: // Все лекарства
                    ShowAllMedicines();
                    break;

                case 1: // Поиск лекарств
                    DiseaseInfo.Visible = false;
                    LblParam1.Text = "Название:";
                    LblParam2.Text = "Форма выпуска:";
                    LblParam3.Text = "Производитель:";
                    LblParam2.Visible = true;
                     TxtSearchParam2.Visible = true;
                    LblParam3.Visible = true;
                    TxtSearchParam3.Visible = true;
                    break;

                case 2: // Лекарства для лечения болезни
                    DiseaseInfo.Visible = true;
                    LblParam1.Text = "Название болезни:";
                    LblParam2.Visible = false;
                     TxtSearchParam2.Visible = false;
                    LblParam3.Visible = false;
                    TxtSearchParam3.Visible = false;

                    // Заполняем подсказки
                    var diseases = dataManager.GetAllDiseases();
                    var autoComplete = new AutoCompleteStringCollection();
                    foreach (var disease in diseases)
                    {
                        autoComplete.Add(disease.Name);
                    }
                    TxtSearchParam1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                    TxtSearchParam1.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    TxtSearchParam1.AutoCompleteCustomSource = autoComplete;
                    break;

                case 3: // Продажи за период
                    DiseaseInfo.Visible = false;
                    monthCalendar.Visible = true;
                   
                    LblParam1.Text = "ID лекарства (опционально):";
                    LblParam2.Visible = false;
                     TxtSearchParam2.Visible = false;
                    LblParam3.Visible = false;
                    TxtSearchParam3.Visible = false;
                    break;

                case 4: // Бронирование
                    CheckAndUpdateExpiredReservations();

                    
                    BtnReserve.Visible = true;
                    BtnExecute.Visible = false;
                    LblParam1.Text = "ID лекарства:";
                    LblParam2.Text = "Имя клиента:";
                    LblParam3.Text = "Количество:";
                    LblParam2.Visible = true;
                    TxtSearchParam2.Visible = true;
                    LblParam3.Visible = true;
                    TxtSearchParam3.Visible = true;
                    break;

                case 5: // Поступление со склада
                    DiseaseInfo.Visible = false;
                    LblParam1.Text = "ID лекарства:";
                    LblParam2.Visible = false;
                     TxtSearchParam2.Visible = false;
                    LblParam3.Visible = false;
                    TxtSearchParam3.Visible = false;
                    break;
            }
        }
        private void CheckAndUpdateExpiredReservations()
        {
            try
            {
                if (dataManager == null) return;

                // Вызываем метод проверки статусов
                dataManager.CheckAndUpdateReservationStatus();

                // Получаем статистику
                string stats = dataManager.GetReservationStatistics();

                // Показываем информационное сообщение
                ShowReservationStatusInfo(stats);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при проверке статуса бронирований: " + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для отображения информации о статусе бронирований
        private void ShowReservationStatusInfo(string statistics)
        {
            try
            {
                // Создаем информационную панель
                Panel infoPanel = new Panel();
                infoPanel.BackColor = Color.LightYellow;
                infoPanel.BorderStyle = BorderStyle.FixedSingle;
                infoPanel.Size = new Size(600, 100);
                infoPanel.Location = new Point(150, 10);

                Label lblInfo = new Label();
                lblInfo.Text = "✓ Статусы бронирований проверены\n" + statistics;
                lblInfo.Location = new Point(10, 10);
                lblInfo.Size = new Size(580, 80);
                lblInfo.Font = new Font("Arial", 9);

                Button btnCloseInfo = new Button();
                btnCloseInfo.Text = "✕";
                btnCloseInfo.Location = new Point(565, 5);
                btnCloseInfo.Size = new Size(25, 25);
                btnCloseInfo.FlatStyle = FlatStyle.Flat;
                btnCloseInfo.Click += (s, ev) =>
                {
                    this.Controls.Remove(infoPanel);
                };

                infoPanel.Controls.Add(lblInfo);
                infoPanel.Controls.Add(btnCloseInfo);

                // Удаляем предыдущую панель, если есть
                foreach (Control control in this.Controls)
                {
                    if (control is Panel && control.Tag != null && control.Tag.ToString() == "ReservationInfoPanel")
                    {
                        this.Controls.Remove(control);
                        break;
                    }
                }

                infoPanel.Tag = "ReservationInfoPanel";
                this.Controls.Add(infoPanel);
                infoPanel.BringToFront();

                // Автоматически скрываем через 10 секунд
                Timer timer = new Timer();
                timer.Interval = 10000; // 10 секунд
                timer.Tick += (s, ev) =>
                {
                    if (this.Controls.Contains(infoPanel))
                    {
                        this.Controls.Remove(infoPanel);
                    }
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();
            }
            catch (Exception)
            {
                // В случае ошибки просто показываем MessageBox
                MessageBox.Show("Статусы бронирований были обновлены.\n\n" + statistics,
                    "Информация о бронированиях", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

     
    
    private void monthCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            // Можно добавить логику при изменении даты
        }

        private void lstResults_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Обработка выбора элемента в списке
        }

        private void PnlSearchParams_Paint(object sender, PaintEventArgs e)
        {
            // Отрисовка панели (по умолчанию)
        }

        private void TxtSearchParam1_TextChanged(object sender, EventArgs e)
        {
            // Обработка изменения текста
        }

        private void  TxtSearchParam2_TextChanged(object sender, EventArgs e)
        {
            // Обработка изменения текста
        }

        private void TxtSearchParam3_TextChanged(object sender, EventArgs e)
        {
            // Обработка изменения текста
        }

        private void LblParam1_Click(object sender, EventArgs e)
        {
            // Обработка клика по метке
        }

        private void LblParam2_Click(object sender, EventArgs e)
        {
            // Обработка клика по метке
        }

        private void LblParam3_Click(object sender, EventArgs e)
        {
            // Обработка клика по метке
        }

        private void BtnExecute_Click(object sender, EventArgs e)
        {
            if (dataManager == null)
            {
                MessageBox.Show("DataManager не инициализирован. Перезапустите приложение.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                switch (CmbFilterType.SelectedIndex)
                {
                    case 0: // Все лекарства
                        ShowAllMedicines();
                        break;

                    case 1: // Поиск лекарств
                        var medicines = dataManager.SearchMedicines(
                            TxtSearchParam1.Text.Trim(),
                             TxtSearchParam2.Text.Trim(),
                            TxtSearchParam3.Text.Trim());
                        lstResults.DataSource = null;
                        lstResults.Items.Clear();
                        foreach (var medicine in medicines)
                        {
                            lstResults.Items.Add(medicine);
                        }
                        if (medicines.Count == 0)
                            MessageBox.Show("Лекарства не найдены", "Результат поиска",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;

                    case 2: // Лекарства для лечения болезни
                        var diseaseName = TxtSearchParam1.Text.Trim();
                        if (!string.IsNullOrEmpty(diseaseName))
                        {
                            var medsForDisease = dataManager.GetMedicinesForDisease(diseaseName);
                            lstResults.DataSource = null;
                            lstResults.Items.Clear();
                            foreach (var medicine in medsForDisease)
                            {
                                lstResults.Items.Add(medicine);
                            }
                            if (medsForDisease.Count == 0)
                                MessageBox.Show(string.Format("Не найдено лекарств для лечения: {0}", diseaseName),
                                    "Результат поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Введите название болезни", "Внимание",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        break;

                    case 3: // Продажи за период
                        var startDate = monthCalendar.SelectionStart;
                        var endDate = monthCalendar.SelectionEnd;

                        int? medicineId = null;
                        if (!string.IsNullOrEmpty( TxtSearchParam1.Text.Trim()))
                        {
                            int parsedId;
                            if (int.TryParse( TxtSearchParam1.Text.Trim(), out parsedId))
                            {
                                medicineId = parsedId;
                            }
                        }

                        var sales = dataManager.GetSalesForPeriod(startDate, endDate, medicineId);
                        lstResults.DataSource = null;
                        lstResults.Items.Clear();
                        foreach (var sale in sales)
                        {
                            lstResults.Items.Add(sale);
                        }

                        // Подсчет общей суммы
                        decimal total = 0;
                        foreach (var sale in sales)
                        {
                            total += sale.Quantity * sale.Price;
                        }
                        if (sales.Count > 0)
                        {
                            lstResults.Items.Add(string.Format("=== ИТОГО: {0} продаж на сумму {1:C} ===",
                                sales.Count, total));
                        }
                        break;

                    case 5: // Поступление со склада
                        int medId;
                        if (int.TryParse( TxtSearchParam1.Text.Trim(), out medId))
                        {
                            var deliveries = dataManager.GetStockDeliveries(medId);
                            lstResults.DataSource = null;
                            lstResults.Items.Clear();
                            foreach (var delivery in deliveries)
                            {
                                lstResults.Items.Add(delivery);
                            }
                            if (deliveries.Count == 0)
                                MessageBox.Show("Поставки не найдены", "Результат поиска",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Введите корректный ID лекарства", "Ошибка",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка выполнения запроса: {0}", ex.Message), "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReserve_Click(object sender, EventArgs e)
        {
            if (dataManager == null)
            {
                MessageBox.Show("DataManager не инициализирован. Перезапустите приложение.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                int medicineId;
                if (!int.TryParse( TxtSearchParam1.Text.Trim(), out medicineId))
                {
                    MessageBox.Show("Введите корректный ID лекарства", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var clientName =  TxtSearchParam2.Text.Trim();
                if (string.IsNullOrEmpty(clientName))
                {
                    MessageBox.Show("Введите имя клиента", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int quantity;
                if (!int.TryParse(TxtSearchParam3.Text.Trim(), out quantity) || quantity <= 0)
                {
                    MessageBox.Show("Введите корректное количество", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var medicine = dataManager.GetMedicineById(medicineId);
                if (medicine == null)
                {
                    MessageBox.Show("Лекарство с указанным ID не найдено", "Ошибка",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Показываем информацию о лекарстве
                var result = MessageBox.Show(
                    string.Format("Забронировать {0} для {1} в количестве {2} шт.?\nЦена за шт.: {3:C}",
                    medicine.Name, clientName, quantity, medicine.Price),
                    "Подтверждение бронирования",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    bool success = dataManager.MakeReservation(medicineId, clientName, quantity);
                    if (success)
                    {
                        MessageBox.Show("Бронирование успешно создано на 3 дня!", "Успех",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Очищаем поля
                        TxtSearchParam1.Clear();
                         TxtSearchParam2.Clear();
                        TxtSearchParam3.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при бронировании: {0}", ex.Message), "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnShowAll_Click(object sender, EventArgs e)
        {
            if (dataManager == null)
            {
                MessageBox.Show("DataManager не инициализирован. Перезапустите приложение.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ShowAllMedicines();
        }

        private void ShowAllMedicines()
        {
            if (dataManager == null)
            {
                MessageBox.Show("DataManager не инициализирован. Перезапустите приложение.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                var allMedicines = dataManager.GetAllMedicines();
                lstResults.DataSource = null;
                lstResults.Items.Clear();
                foreach (var medicine in allMedicines)
                {
                    lstResults.Items.Add(medicine);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при загрузке лекарств: {0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DiseaseInfo_Click(object sender, EventArgs e)
        {
            if (dataManager == null)
            {
                MessageBox.Show("DataManager не инициализирован. Перезапустите приложение.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Получаем список всех болезней
                var diseases = dataManager.GetAllDiseases();

                if (diseases == null || diseases.Count == 0)
                {
                    MessageBox.Show("Список болезней пуст.",
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Создаем сообщение со списком болезней
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Консультации возможны по болезням:");
                sb.AppendLine();

                foreach (var disease in diseases)
                {
                    sb.AppendLine(string.Format("• {0} - {1}",
                        disease.Name, disease.Description));
                }

                sb.AppendLine();
                sb.AppendLine("Для поиска лекарств введите название болезни в поле выше.");

                // Показываем MessageBox со списком болезней
                MessageBox.Show(sb.ToString(),
                    "Справочник болезней",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при загрузке списка болезней: {0}", ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void checkAdmin_CheckedChanged(object sender, EventArgs e)
        {
            Booking.Visible = checkAdmin.Checked;
            Reset.Visible = checkAdmin.Checked;
        }

        private void Booking_Click(object sender, EventArgs e)
        {
            string filePath = "Reservation.csv";
            System.Diagnostics.Process.Start(filePath);
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Reset_Click(object sender, EventArgs e)
        {
            if (dataManager == null)
            {
                MessageBox.Show("DataManager не инициализирован. Перезапустите приложение.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Показываем диалог с выбором действия
            Form resetDialog = new Form();
            resetDialog.Text = "Управление бронированиями";
            resetDialog.Size = new Size(500, 350);
            resetDialog.StartPosition = FormStartPosition.CenterParent;

            Label lblTitle = new Label();
            lblTitle.Text = "ВЫБЕРИТЕ ДЕЙСТВИЕ:";
            lblTitle.Location = new Point(20, 20);
            lblTitle.Size = new Size(450, 30);
            lblTitle.Font = new Font("Arial", 12, FontStyle.Bold);
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;

            Button btnSoftReset = new Button();
            btnSoftReset.Text = "1. Отменить активные бронирования\n(вернуть лекарства, статус → Cancelled)";
            btnSoftReset.Location = new Point(50, 70);
            btnSoftReset.Size = new Size(400, 60);
            btnSoftReset.TextAlign = ContentAlignment.MiddleCenter;
            btnSoftReset.Click += (s, ev) =>
            {
                ExecuteSoftReset();
                resetDialog.Close();
            };

            Button btnHardReset = new Button();
            btnHardReset.Text = "2. Очистить ВСЕ бронирования\n(удалить все записи, файл станет пустым)";
            btnHardReset.Location = new Point(50, 140);
            btnHardReset.Size = new Size(400, 60);
            btnHardReset.TextAlign = ContentAlignment.MiddleCenter;
            btnHardReset.BackColor = Color.LightCoral;
            btnHardReset.Click += (s, ev) =>
            {
                ExecuteHardReset();
                resetDialog.Close();
            };

            Button btnBackup = new Button();
            btnBackup.Text = "3. Создать резервную копию данных";
            btnBackup.Location = new Point(50, 210);
            btnBackup.Size = new Size(400, 40);
            btnBackup.Click += (s, ev) =>
            {
                CreateBackup();
                resetDialog.Close();
            };

            Button btnCancel = new Button();
            btnCancel.Text = "Отмена";
            btnCancel.Location = new Point(200, 260);
            btnCancel.Size = new Size(100, 35);
            btnCancel.Click += (s, ev) => resetDialog.Close();

            resetDialog.Controls.AddRange(new Control[] {
            lblTitle, btnSoftReset, btnHardReset, btnBackup, btnCancel
        });

            resetDialog.ShowDialog();
        }

        // Мягкий сброс (только активные брони)
        private void ExecuteSoftReset()
        {
            DialogResult confirm = MessageBox.Show(
                "Вы уверены, что хотите отменить ВСЕ активные бронирования?\n\n" +
                "✓ Лекарства будут возвращены в общее количество\n" +
                "✓ Статус броней изменится на 'Cancelled'\n" +
                "✓ История бронирований сохранится\n\n" +
                "Эта операция необратима!",
                "Подтверждение отмены бронирований",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm == DialogResult.Yes)
            {
                string result = dataManager.ResetAllReservations();

                MessageBox.Show(result,
                    "Результат операции",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // Обновляем отображение, если нужно
                if (CmbFilterType.SelectedIndex == 4) // Если в разделе бронирования
                {
                    // Можно обновить статистику или очистить поля
                    TxtSearchParam1.Clear();
                    TxtSearchParam2.Clear();
                    TxtSearchParam3.Clear();
                }
            }
        }

        // Полная очистка (все брони)
        private void ExecuteHardReset()
        {
            DialogResult confirm = MessageBox.Show(
                "⚠️  ВНИМАНИЕ! КРИТИЧЕСКАЯ ОПЕРАЦИЯ!\n\n" +
                "Вы собираетесь УДАЛИТЬ ВСЕ записи о бронированиях!\n\n" +
                "✓ Все бронирования будут удалены\n" +
                "✓ Активные брони вернут лекарства\n" +
                "✓ Файл Reservation.csv станет пустым\n" +
                "✓ История бронирований будет утеряна\n\n" +
                "Рекомендуется создать резервную копию перед выполнением!\n\n" +
                "ПОДТВЕРДИТЕ УДАЛЕНИЕ ВСЕХ БРОНИРОВАНИЙ:",
                "КРИТИЧЕСКОЕ ПРЕДУПРЕЖДЕНИЕ",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Error);

            if (confirm == DialogResult.Yes)
            {
                string result = dataManager.ClearAllReservations();

                MessageBox.Show(result,
                    "Результат очистки",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        // Создание резервной копии
        private void CreateBackup()
        {
            string result = dataManager.CreateBackup();
            MessageBox.Show(result,
                "Резервное копирование",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // ... остальные методы без изменений ...
    }
}
    
 