using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LibrarsStraightLine
{
    public static class BeamLine
    {
        public static string Intersection_Of_Axes_BeamLine(string input)
        {
            string[] pointParts = input.Split(' ');

            if (pointParts.Length != 4)
            {
                return pointParts.Length < 4 ? "Не достаточно данных" : "Ошибка в обработке";
            }

            double x1 = Convert.ToDouble(pointParts[0]);
            double y1 = Convert.ToDouble(pointParts[1]);
            double x2 = Convert.ToDouble(pointParts[2]);
            double y2 = Convert.ToDouble(pointParts[3]);

            // Направляющий вектор луча
            double dx = x2 - x1;
            double dy = y2 - y1;

            bool intersectsOX = false;
            bool intersectsOY = false;

            // Проверка пересечения с осью OX (y = 0)
            if (Math.Abs(dy) > 1e-10) // Луч не горизонтальный
            {
                double t_ox = -y1 / dy; // Параметр t, при котором y = 0
                if (t_ox >= 0) // Пересечение происходит в направлении луча (t ≥ 0)
                {
                    intersectsOX = true;
                }
            }
            else if (Math.Abs(y1) < 1e-10) // Луч горизонтальный и лежит на оси OX
            {
                intersectsOX = true;
            }

            // Проверка пересечения с осью OY (x = 0)
            if (Math.Abs(dx) > 1e-10) // Луч не вертикальный
            {
                double t_oy = -x1 / dx; // Параметр t, при котором x = 0
                if (t_oy >= 0) // Пересечение происходит в направлении луча (t ≥ 0)
                {
                    intersectsOY = true;
                }
            }
            else if (Math.Abs(x1) < 1e-10) // Луч вертикальный и лежит на оси OY
            {
                intersectsOY = true;
            }

            // Определение результата
            if (intersectsOX && intersectsOY)
            {
                return "Пересекает оси OY и OX";
            }
            else if (intersectsOX)
            {
                return "Пересекает ось OX";
            }
            else if (intersectsOY)
            {
                return "Пересекает ось OY";
            }
            else
            {
                return "Не пересекает ни одну ось";
            }
        }
    }
    }
