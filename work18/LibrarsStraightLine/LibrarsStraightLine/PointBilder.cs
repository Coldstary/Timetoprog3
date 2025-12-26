using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LibrarsStraightLine
{
    // Статический класс для работы с геометрическими преобразованиями и вычислениями
    public static class PointBilder
    {
        // Метод для преобразования строки координат в форматированный вид точек
        public static string ConvertToPointString(string input)
        {
            // Разделение входной строки на части по пробелам
            string[] pointParts = input.Split(' ');

            // Проверка наличия всех четырех координат (две точки: x1, y1, x2, y2)
            if (pointParts.Length == 4)
            {
                // Форматирование координат в виде двух точек: (x1, y1), (x2, y2)
                return $"({pointParts[0]}, {pointParts[1]}),({pointParts[2]}, {pointParts[3]})";
            }
            // Проверка на недостаточное количество данных
            else if (pointParts.Length < 4)
            {
                // Сообщение об ошибке с указанием требуемого формата
                return "(не достаточно данных. Введите 4 значения: x1, y1, x2,y2)";
            }

            // Возврат пустой строки в случае других ситуаций
            return string.Empty;
        }

        // Метод для вычисления угла между отрезком и осью OX
        public static string ConvertToPointCornerOX(string input)
        {
            // Разделение входной строки на части по пробелам
            string[] pointParts = input.Split(' ');

            // Проверка наличия координат (<=4 для обработки ровно 4 координат)
            if (pointParts.Length <= 4)
            {
                // Вычисление разности координат по оси X с использованием модуля
                double AbesX = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[0])) - Math.Abs(Convert.ToInt32(pointParts[2])));
                // Вычисление разности координат по оси Y с использованием модуля
                double AbesY = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[1])) - Math.Abs(Convert.ToInt32(pointParts[3])));
                // Вычисление угла через арктангенс отношения вертикальной и горизонтальной составляющих
                double CornerOX = Math.Atan(AbesY / AbesX);

                // Возврат результата в радианах
                return $"Угол от оси X: {CornerOX} радиан";
            }
            // Проверка на недостаточное количество данных
            else if (pointParts.Length < 4)
            {
                // Сообщение об ошибке с указанием требуемого формата
                return "(не достаточно данных.Введите значения x1 y1 x2 y2 )";
            }

            // Сообщение об общей ошибке
            return "Ошибка";
        }

        // Метод для проверки параллельности отрезка осям координат
        public static string ConvertToPointParal(string input)
        {
            // Разделение входной строки на части по пробелам
            string[] pointParts = input.Split(' ');

            // Проверка наличия всех четырех координат
            if (pointParts.Length == 4)
            {
                // Проверка параллельности оси Y: равны ли x-координаты двух точек
                if ((Convert.ToInt32(pointParts[0])) == (Convert.ToInt32(pointParts[2])))
                {
                    return $"Построение паралельно оси Y";
                }
                // Проверка параллельности оси X: равны ли y-координаты двух точек
                else if ((Convert.ToInt32(pointParts[1])) == (Convert.ToInt32(pointParts[3])))
                {
                    return $"Построение паралельно оси X";
                }
                else
                {
                    // Если ни одно из условий не выполняется
                    return $"Парралельность остутствует";
                }
            }
            // Проверка на недостаточное количество данных
            else if (pointParts.Length < 4)
            {
                return "(не достаточно данных)";
            }

            // Сообщение об ошибке обработки
            return "Ошибка в обработке";
        }

        // Метод для вычисления  Угла от оси X
        public static string ConvertToPointLenth(string input)
        {
            // Разделение входной строки на части по пробелам
            string[] pointParts = input.Split(' ');

            // Проверка наличия всех четырех координат
            if (pointParts.Length == 4)
            {
                // Вычисление разности координат по оси X с использованием модуля
                double AbesX = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[0])) - Math.Abs(Convert.ToInt32(pointParts[2])));
                // Вычисление разности координат по оси Y с использованием модуля
                double AbesY = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[1])) - Math.Abs(Convert.ToInt32(pointParts[3])));
                // Вычисление длины отрезка по теореме Пифагора
                double LenthOX = Math.Atan(AbesY / AbesX );

                // Угол от оси X
                return $"Угол от оси X: {LenthOX} радиан";
            }
            // Проверка на недостаточное количество данных
            else if (pointParts.Length < 4)
            {
                return "(не достаточно данных)";
            }

            // Сообщение об ошибке обработки
            return "Ошибка в обработке";
        }
    }
}

