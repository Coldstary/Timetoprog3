using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LibrarsStraightLine
{
    public static class PointBilder
    {
        public static string ConvertToPointString(string input)
        {

            string[] pointParts = input.Split(' ');

            if (pointParts.Length == 4)
            {
                // Форматируем как точку (x, y)
                return $"({pointParts[0]}, {pointParts[1]}),({pointParts[2]}, {pointParts[3]})";
            }
            else if (pointParts.Length < 4)
            {
                return "(че происходит)";
            }

            return string.Empty;
        }
        public static string ConvertToPointCornerOX(string input)
        {

            string[] pointParts = input.Split(' ');

            if (pointParts.Length == 4)
            {
                // Форматируем как точку (x, y)
                double AbesX = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[0])) - Math.Abs(Convert.ToInt32(pointParts[2])));
                double AbesY = Math.Abs(Math.Abs(Convert.ToInt32(pointParts[1])) - Math.Abs(Convert.ToInt32(pointParts[3])));
                double CornerOX = Math.Atan(AbesY / AbesX);

                return $"Угол от оси X: {CornerOX}";
            }
            else if (pointParts.Length < 4)
            {
                return "(че происходит2)";
            }

            return "скип1";
        }
        public static string ConvertToPointParal(string input)
        {

            string[] pointParts = input.Split(' ');

            if (pointParts.Length == 4)
            {
                if (Math.Abs(Convert.ToInt32(pointParts[0])) == Math.Abs(Convert.ToInt32(pointParts[2])))
                {
                    return $"Построение паралельно оси X";
                }
                else if(Math.Abs(Convert.ToInt32(pointParts[1])) == Math.Abs(Convert.ToInt32(pointParts[3])))
                {
                    return $"Построение паралельно оси Y";
                }
                else
                { return $"Парралельность остутствует"; }
            }
            else if (pointParts.Length < 4)
            {
                return "(че происходит2)";
            }

            return "скип2";
        }

    }
    
}



