using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CsvWork21new
{
    public class DataManager
    {
        private List<Medicine> medicines = new List<Medicine>();
        private List<Stock> stocks = new List<Stock>();
        private List<PharmacyStock> pharmacyStocks = new List<PharmacyStock>();
        private List<Reservation> reservations = new List<Reservation>();
        private List<Sale> sales = new List<Sale>();
        private List<Disease> diseases = new List<Disease>();
        private List<MedicineDisease> medicineDiseases = new List<MedicineDisease>();

        public DataManager()
        {
            try
            {
                LoadAllData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных: " + ex.Message,
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAllData()
        {
            medicines = LoadCSV<Medicine>("Medicine.csv");
            stocks = LoadCSV<Stock>("Stock.csv");
            pharmacyStocks = LoadCSV<PharmacyStock>("PharmacyStock.csv");
            reservations = LoadCSV<Reservation>("Reservation.csv");
            sales = LoadCSV<Sale>("Sale.csv");
            diseases = LoadCSV<Disease>("Disease.csv");
            medicineDiseases = LoadCSV<MedicineDisease>("MedicineDisease.csv");
        }

        private List<T> LoadCSV<T>(string filename) where T : class, new()
        {
            var list = new List<T>();

            if (!File.Exists(filename))
            {
                MessageBox.Show(string.Format("Файл {0} не найден!", filename),
                    "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return list;
            }

            try
            {
                var lines = File.ReadAllLines(filename, Encoding.UTF8);
                if (lines.Length == 0) return list;

                // Убираем BOM из первой строки если есть
                lines[0] = lines[0].Replace("\uFEFF", "");

                var properties = typeof(T).GetProperties();

                for (int i = 1; i < lines.Length; i++)
                {
                    var values = lines[i].Split(';');
                    if (values.Length == 0) continue;

                    var obj = new T();

                    for (int j = 0; j < properties.Length && j < values.Length; j++)
                    {
                        var prop = properties[j];
                        var value = values[j];

                        if (string.IsNullOrEmpty(value)) continue;

                        try
                        {
                            if (prop.PropertyType == typeof(int))
                            {
                                int intValue;
                                if (int.TryParse(value, out intValue))
                                {
                                    prop.SetValue(obj, intValue, null);
                                }
                            }
                            else if (prop.PropertyType == typeof(decimal))
                            {
                                decimal decimalValue;
                                if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimalValue))
                                {
                                    prop.SetValue(obj, decimalValue, null);
                                }
                            }
                            else if (prop.PropertyType == typeof(DateTime))
                            {
                                DateTime dateValue;
                                if (DateTime.TryParse(value, out dateValue))
                                {
                                    prop.SetValue(obj, dateValue, null);
                                }
                            }
                            else if (prop.PropertyType == typeof(string))
                            {
                                prop.SetValue(obj, value, null);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(string.Format("Ошибка при парсинге значения {0} для свойства {1}: {2}",
                                value, prop.Name, ex.Message));
                        }
                    }

                    list.Add(obj);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при загрузке файла {0}: {1}", filename, ex.Message),
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return list;
        }
        // Методы для запросов LINQ

        public List<Medicine> GetAllMedicines()
        {
            return medicines;
        }

        public List<Medicine> SearchMedicines(string name, string form, string manufacturer)
        {
            var query = medicines.AsQueryable();

            if (!string.IsNullOrEmpty(name))
            {
                string nameLower = name.ToLower();
                query = query.Where(m => m.Name.ToLower().Contains(nameLower));
            }

            if (!string.IsNullOrEmpty(form))
            {
                string formLower = form.ToLower();
                query = query.Where(m => m.Form.ToLower().Contains(formLower));
            }

            if (!string.IsNullOrEmpty(manufacturer))
            {
                string manufacturerLower = manufacturer.ToLower();
                query = query.Where(m => m.Manufacturer.ToLower().Contains(manufacturerLower));
            }

            return query.ToList();
        }

        public List<Medicine> GetMedicinesForDisease(string diseaseName)
        {
            // Используем ToLower для регистронезависимого сравнения
            var disease = diseases.FirstOrDefault(d =>
                d.Name.ToLower() == diseaseName.ToLower());

            if (disease == null)
                return new List<Medicine>();

            var medicineIds = medicineDiseases
                .Where(md => md.DiseaseId == disease.Id)
                .Select(md => md.MedicineId)
                .ToList();

            return medicines
                .Where(m => medicineIds.Contains(m.Id))
                .ToList();
        }

        public List<Sale> GetSalesForPeriod(DateTime startDate, DateTime endDate, int? medicineId = null)
        {
            var query = sales.Where(s => s.SaleDate >= startDate && s.SaleDate <= endDate);

            if (medicineId.HasValue)
                query = query.Where(s => s.MedicineId == medicineId.Value);

            return query.ToList();
        }

        public bool MakeReservation(int medicineId, string clientName, int quantity)
        {
            var medicine = medicines.FirstOrDefault(m => m.Id == medicineId);
            if (medicine == null)
            {
                System.Windows.Forms.MessageBox.Show("Лекарство не найдено!");
                return false;
            }

            var pharmacyStock = pharmacyStocks.FirstOrDefault(ps => ps.MedicineId == medicineId);
            if (pharmacyStock == null || pharmacyStock.Quantity < quantity)
            {
                System.Windows.Forms.MessageBox.Show("Недостаточно лекарств в наличии!");
                return false;
            }

            // Проверяем, не истек ли срок годности
            if (medicine.ExpirationDate < DateTime.Now)
            {
                System.Windows.Forms.MessageBox.Show("Срок годности лекарства истек!");
                return false;
            }

            // Находим максимальный ID для нового бронирования
            int newId = 1;
            if (reservations.Count > 0)
            {
                newId = reservations.Max(r => r.Id) + 1;
            }

            var reservation = new Reservation
            {
                Id = newId,
                MedicineId = medicineId,
                ClientName = clientName,
                ReservationDate = DateTime.Now,
                Quantity = quantity,
                Status = "Active"
            };

            reservations.Add(reservation);

            // Уменьшаем количество в наличии
            pharmacyStock.Quantity -= quantity;

            SaveReservations();
            SavePharmacyStock();

            return true;
        }

        public List<Stock> GetStockDeliveries(int medicineId)
        {
            return stocks
                .Where(s => s.MedicineId == medicineId)
                .OrderByDescending(s => s.DeliveryDate)
                .ToList();
        }

        public List<Disease> GetAllDiseases()
        {
            return diseases;
        }

        public Medicine GetMedicineById(int id)
        {
            return medicines.FirstOrDefault(m => m.Id == id);
        }

        private void SaveReservations()
        {
            SaveCSV("Reservation.csv", reservations);
        }

        private void SavePharmacyStock()
        {
            SaveCSV("PharmacyStock.csv", pharmacyStocks);
        }

        private void SaveCSV<T>(string filename, List<T> data)
        {
            try
            {
                var properties = typeof(T).GetProperties();
                var lines = new List<string>();

                // Заголовок
                var header = new List<string>();
                foreach (var prop in properties)
                {
                    header.Add(prop.Name);
                }
                lines.Add(string.Join(";", header.ToArray()));

                // Данные
                foreach (var item in data)
                {
                    var rowValues = new List<string>();
                    foreach (var prop in properties)
                    {
                        var value = prop.GetValue(item, null);
                        if (value is DateTime)
                        {
                            rowValues.Add(((DateTime)value).ToString("yyyy-MM-dd"));
                        }
                        else if (value != null)
                        {
                            rowValues.Add(value.ToString());
                        }
                        else
                        {
                            rowValues.Add("");
                        }
                    }
                    lines.Add(string.Join(";", rowValues.ToArray()));
                }

                File.WriteAllLines(filename, lines.ToArray(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(string.Format("Ошибка при сохранении файла {0}: {1}", filename, ex.Message));
            }
        }
        public void CheckAndUpdateReservationStatus()
        {
            try
            {
                bool changesMade = false;
                DateTime currentDate = DateTime.Now;

                foreach (var reservation in reservations.ToList()) // Используем ToList() для копирования коллекции
                {
                    // Проверяем, прошло ли более 3 дней с даты бронирования
                    TimeSpan timeSinceReservation = currentDate - reservation.ReservationDate;

                    if (reservation.Status == "Active" && timeSinceReservation.TotalDays > 3)
                    {
                        // Меняем статус на NotActive
                        reservation.Status = "NotActive";

                        // Возвращаем количество лекарств в наличие аптеки
                        var pharmacyStock = pharmacyStocks.FirstOrDefault(ps => ps.MedicineId == reservation.MedicineId);
                        if (pharmacyStock != null)
                        {
                            pharmacyStock.Quantity += reservation.Quantity;
                            changesMade = true;
                        }
                    }
                }

                // Если были изменения, сохраняем их в файлы
                if (changesMade)
                {
                    SaveReservations();
                    SavePharmacyStock();

                    // Записываем в лог
                    LogReservationUpdate();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ошибка при обновлении статуса бронирований: " + ex.Message);
            }
        }

        // Метод для логирования обновлений статуса
        private void LogReservationUpdate()
        {
            try
            {
                string logFile = "ReservationLog.txt";
                string logEntry = string.Format("[{0}] Обновлены статусы бронирований: устаревшие брони переведены в NotActive",
                    DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"));

                File.AppendAllText(logFile, logEntry + Environment.NewLine);
            }
            catch (Exception)
            {
                // Игнорируем ошибки логирования
            }
        }

        // Метод для получения статистики по бронированиям
        public string GetReservationStatistics()
        {
            int totalReservations = reservations.Count;
            int activeReservations = reservations.Count(r => r.Status == "Active");
            int notActiveReservations = reservations.Count(r => r.Status == "NotActive");
            int expiredReservations = reservations.Count(r =>
                r.Status == "Active" && (DateTime.Now - r.ReservationDate).TotalDays > 3);

            return string.Format("Всего бронирований: {0}\n" +
                                "Активных: {1}\n" +
                                "Неактивных: {2}\n" +
                                "Просроченных (требуют обновления): {3}",
                                totalReservations, activeReservations, notActiveReservations, expiredReservations);
        }
        public string ResetAllReservations()
        {
            StringBuilder result = new StringBuilder();
            int activeReservationsCount = 0;
            int returnedQuantity = 0;

            try
            {
                // 1. Находим все активные бронирования
                var activeReservations = reservations.Where(r => r.Status == "Active").ToList();
                activeReservationsCount = activeReservations.Count;

                if (activeReservationsCount == 0)
                {
                    return "Активных бронирований нет. Сброс не требуется.";
                }

                // 2. Возвращаем лекарства в PharmacyStock
                foreach (var reservation in activeReservations)
                {
                    var pharmacyStock = pharmacyStocks.FirstOrDefault(ps => ps.MedicineId == reservation.MedicineId);
                    if (pharmacyStock != null)
                    {
                        pharmacyStock.Quantity += reservation.Quantity;
                        returnedQuantity += reservation.Quantity;

                        // Получаем название лекарства для отчета
                        var medicine = medicines.FirstOrDefault(m => m.Id == reservation.MedicineId);
                        string medicineName = medicine != null ? medicine.Name : $"ID:{reservation.MedicineId}";

                        result.AppendLine(string.Format("• {0}: возвращено {1} единиц",
                            medicineName, reservation.Quantity));
                    }

                    // 3. Меняем статус на "Cancelled" (отменено сбросом)
                    reservation.Status = "Cancelled";
                }

                // 4. Сохраняем изменения
                SaveReservations();
                SavePharmacyStock();

                // 5. Логируем операцию
                LogResetOperation(activeReservationsCount, returnedQuantity);

                return string.Format("СБРОС БРОНИРОВАНИЙ ВЫПОЛНЕН УСПЕШНО!\n\n" +
                                   "Отменено бронирований: {0}\n" +
                                   "Возвращено лекарств: {1} единиц\n\n" +
                                   "Детали возврата:\n{2}",
                                   activeReservationsCount, returnedQuantity, result.ToString());
            }
            catch (Exception ex)
            {
                return string.Format("Ошибка при сбросе бронирований: {0}", ex.Message);
            }
        }

        // Метод для полной очистки всех бронирований (удаление всех записей)
        public string ClearAllReservations()
        {
            try
            {
                int totalReservations = reservations.Count;
                int activeReservations = reservations.Count(r => r.Status == "Active");

                // Сначала возвращаем активные бронирования
                string resetResult = ResetAllReservations();

                // Затем очищаем весь список бронирований
                reservations.Clear();

                // Сохраняем пустой файл
                SaveReservations();

                return string.Format("ВСЕ БРОНИРОВАНИЯ ОЧИЩЕНЫ!\n\n" +
                                   "Всего удалено записей: {0}\n" +
                                   "Из них активных: {1}\n\n" +
                                   "Результат возврата лекарств:\n{2}",
                                   totalReservations, activeReservations, resetResult);
            }
            catch (Exception ex)
            {
                return string.Format("Ошибка при очистке бронирований: {0}", ex.Message);
            }
        }

        // Метод для логирования операции сброса
        private void LogResetOperation(int reservationsCount, int returnedQuantity)
        {
            try
            {
                string logFile = "ResetLog.txt";
                string logEntry = string.Format("[{0}] СБРОС БРОНИРОВАНИЙ\n" +
                                              "Отменено бронирований: {1}\n" +
                                              "Возвращено единиц: {2}\n" +
                                              "Пользователь: {3}\n" +
                                              "--------------------------------\n",
                    DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"),
                    reservationsCount,
                    returnedQuantity,
                    Environment.UserName);

                File.AppendAllText(logFile, logEntry, Encoding.UTF8);
            }
            catch (Exception)
            {
                // Игнорируем ошибки логирования
            }
        }

        // Метод для восстановления PharmacyStock из резервной копии
        public string RestorePharmacyStockFromBackup(string backupDate = "")
        {
            try
            {
                string backupDir = "Backups";
                if (!Directory.Exists(backupDir))
                {
                    return "Директория резервных копий не найдена.";
                }

                // Ищем файлы бэкапов PharmacyStock
                var backupFiles = Directory.GetFiles(backupDir, "PharmacyStock_*.bak")
                                          .OrderByDescending(f => f)
                                          .ToList();

                if (backupFiles.Count == 0)
                {
                    return "Резервные копии не найдены.";
                }

                string backupFile;
                if (string.IsNullOrEmpty(backupDate))
                {
                    // Берем последнюю резервную копию
                    backupFile = backupFiles.First();
                }
                else
                {
                    // Ищем по дате
                    backupFile = backupFiles.FirstOrDefault(f => f.Contains(backupDate));
                    if (backupFile == null)
                    {
                        return string.Format("Резервная копия за дату {0} не найдена.", backupDate);
                    }
                }

                // Копируем резервную копию в основной файл
                File.Copy(backupFile, "PharmacyStock.csv", true);

                // Перезагружаем данные
                pharmacyStocks = LoadCSV<PharmacyStock>("PharmacyStock.csv");

                return string.Format("PharmacyStock успешно восстановлен из резервной копии:\n{0}",
                    Path.GetFileName(backupFile));
            }
            catch (Exception ex)
            {
                return string.Format("Ошибка при восстановлении: {0}", ex.Message);
            }
        }

        // Метод для создания резервной копии текущего состояния
        public string CreateBackup()
        {
            try
            {
                string backupDir = "Backups";
                if (!Directory.Exists(backupDir))
                {
                    Directory.CreateDirectory(backupDir);
                }

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string backupFile = Path.Combine(backupDir,
                    string.Format("FullBackup_{0}.zip", timestamp));

                // Создаем список файлов для архивации
                string[] filesToBackup = {
                "Medicine.csv",
                "Stock.csv",
                "PharmacyStock.csv",
                "Reservation.csv",
                "Sale.csv",
                "Disease.csv",
                "MedicineDisease.csv"
            };

                // Просто копируем файлы (вместо архивации для простоты)
                foreach (string file in filesToBackup)
                {
                    if (File.Exists(file))
                    {
                        string backupCopy = Path.Combine(backupDir,
                            string.Format("{0}_{1}.bak",
                                Path.GetFileNameWithoutExtension(file),
                                timestamp));
                        File.Copy(file, backupCopy, true);
                    }
                }

                return string.Format("Резервная копия создана успешно.\nВремя: {0}",
                    DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"));
            }
            catch (Exception ex)
            {
                return string.Format("Ошибка при создании резервной копии: {0}", ex.Message);
            }
        }
    }
}
