using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CsvWork21new
{
    // Класс лекарства
    public class Medicine
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Form { get; set; }
        public DateTime ExpirationDate { get; set; }
        public string Annotation { get; set; }
        public decimal Price { get; set; }
        public string Manufacturer { get; set; }

        public override string ToString()
        {
            return $"{Name} ({Form}) - {Manufacturer}, Цена: {Price:C}, Срок годности: {ExpirationDate:dd.MM.yyyy}";
        }
    }

    // Класс склад
    public class Stock
    {
        public int Id { get; set; }
        public int MedicineId { get; set; }
        public int Quantity { get; set; }
        public DateTime DeliveryDate { get; set; }

        public override string ToString()
        {
            return $"Поставка от {DeliveryDate:dd.MM.yyyy}, Количество: {Quantity}";
        }
    }

    // Класс наличие в аптеке
    public class PharmacyStock
    {
        public int Id { get; set; }
        public int MedicineId { get; set; }
        public int Quantity { get; set; }
        public DateTime LastDeliveryDate { get; set; }

        public override string ToString()
        {
            return $"В наличии: {Quantity} шт., Последняя поставка: {LastDeliveryDate:dd.MM.yyyy}";
        }
    }

    // Класс бронирование
    public class Reservation
    {
        public int Id { get; set; }
        public int MedicineId { get; set; }
        public string ClientName { get; set; }
        public DateTime ReservationDate { get; set; }
        public int Quantity { get; set; }
        public string Status { get; set; }

        public override string ToString()
        {
            return $"{ClientName}, Количество: {Quantity}, Дата: {ReservationDate:dd.MM.yyyy}, Статус: {Status}";
        }
    }

    // Класс продажа
    public class Sale
    {
        public int Id { get; set; }
        public int MedicineId { get; set; }
        public DateTime SaleDate { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public string ClientName { get; set; }

        public override string ToString()
        {
            return $"{SaleDate:dd.MM.yyyy}, Количество: {Quantity}, Сумма: {Quantity * Price:C}, Клиент: {ClientName}";
        }
    }

    // Класс болезнь
    public class Disease
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }

        public override string ToString()
        {
            return $"{Name} - {Description}";
        }
    }

    // Класс связь лекарство-болезнь
    public class MedicineDisease
    {
        public int Id { get; set; }
        public int MedicineId { get; set; }
        public int DiseaseId { get; set; }
    }
}