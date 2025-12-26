using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;  // Не используется в текущем коде
using System.Text;

namespace LibrarsStraightLine
{
    // Статический класс для работы с отрезками (ограниченными линиями)
    public static class Limited_Line
    {
        // Метод определения пересечения отрезка с осями координат
        public static string Intersection_Of_Axes_Limited_Line(string input)
        {
            int Axes = 0;  // Счетчик пересеченных осей (0, 1 или 2)
            string[] pointParts = input.Split(' ');  // Разделение строки на координаты

            // Проверка наличия всех четырех координат (две точки)
            if (pointParts.Length == 4)
            {
                // Проверка пересечения с осью OX:
                // Логика: если одна точка имеет x <= 0, а другая x >= 0, отрезок пересекает ось OY
                // ВНИМАНИЕ: в тексте сообщения указано "ось OX", но проверяются x-координаты (ось OY)
                if ((((Convert.ToInt32(pointParts[0])) <= 0) & (Convert.ToInt32(pointParts[2])) >= 0) ||
                    (((Convert.ToInt32(pointParts[0])) >= 0) & (Convert.ToInt32(pointParts[2])) <= 0))
                {
                    Axes += 1;  // Увеличение счетчика пересеченных осей
                    return $"Пересекает ось OX";  // Возврат результата (ось OY на самом деле)
                }
                // Проверка пересечения с осью OY:
                // Логика: если одна точка имеет y <= 0, а другая y >= 0, отрезок пересекает ось OX
                // ВНИМАНИЕ: в тексте сообщения указано "ось OY", но проверяются y-координаты (ось OX)
                else if ((((Convert.ToInt32(pointParts[1])) <= 0) & (Convert.ToInt32(pointParts[3])) >= 0) ||
                         (((Convert.ToInt32(pointParts[1])) >= 0) & (Convert.ToInt32(pointParts[3])) <= 0))
                {
                    Axes += 1;  // Увеличение счетчика пересеченных осей
                    return $"Пересекает ось OY";  // Возврат результата (ось OX на самом деле)
                }
                // Проверка на пересечение обеих осей (это условие никогда не выполнится из-за ранних return)
                else if (Axes == 2)
                {
                    return $"Пересекает оси OY и OX";
                }
                else
                {
                    // Если ни одно из условий не выполнилось
                    return ("Не пересекает ни одну ось");
                }
            }
            // Проверка на недостаточное количество данных
            else if (pointParts.Length < 4)
            {
                return "(не достаточно данных)";
            }

            // Возврат при других ошибках
            return "Ошибка в обработке";
        }

        // Метод вычисления длины отрезка
        public static string Length_Limited_Line(string input)
        {
            string[] pointParts = input.Split(' ');  // Разделение строки на координаты

            // Проверка наличия координат (<=4 для обработки ровно 4 координат)
            if (pointParts.Length <= 4)
            {
                // Вычисление разности координат по оси X (абсолютное значение)
                double AbesX = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[0])) - Math.Abs(Convert.ToInt32(pointParts[2])));
                // Вычисление разности координат по оси Y (абсолютное значение)
                double AbesY = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[1])) - Math.Abs(Convert.ToInt32(pointParts[3])));
                // Вычисление длины отрезка по теореме Пифагора
                double Length = Math.Sqrt(AbesY * AbesY + AbesX * AbesX);

                // Возврат вычисленной длины отрезка
                return $"Длинна отрезка: {Length} ";
            }
            // Проверка на недостаточное количество данных
            else if (pointParts.Length < 4)
            {
                // Сообщение об ошибке с указанием требуемого формата
                return "(не достаточно данных. Введите 4 значения: x1, y1, x2,y2 через пробел и повторите попытку)";
            }

            // Сообщение об общей ошибке
            return "Ошибка";
        }
    }
}