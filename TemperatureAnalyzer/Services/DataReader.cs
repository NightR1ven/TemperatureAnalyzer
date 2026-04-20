using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TemperatureAnalyzer.Models;
using System.IO;

namespace TemperatureAnalyzer.Services
{
    /// <summary>
    /// Сервис для чтения исходного файла ISH.txt
    /// </summary>
    public static class DataReader
    {
        public static List<DataPoint> ReadData(string filePath, out double pH)
        {
            var points = new List<DataPoint>();
            pH = 0.0;

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Файл {filePath} не найден.");

            string[] lines = File.ReadAllLines(filePath, System.Text.Encoding.Default);
            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(',');
                // Последняя строка – атмосферное давление (одно число)
                if (parts.Length == 1 && double.TryParse(parts[0], System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out pH))
                {
                    break;
                }

                if (parts.Length < 15) continue; // ожидаем 15 полей

                try
                {
                    var dp = new DataPoint
                    {
                        Field0 = ParseDouble(parts[0]),
                        Field1 = ParseDouble(parts[1]),
                        Field2 = ParseDouble(parts[2]),
                        Field3 = ParseDouble(parts[3]),
                        Field4 = ParseDouble(parts[4]),
                        Field5 = ParseDouble(parts[5]),
                        Field6 = ParseDouble(parts[6]),
                        Field7 = ParseDouble(parts[7]),
                        Field8 = ParseDouble(parts[8]),
                        t4 = new double[6]
                        {
                            ParseDouble(parts[8]),  // T4_1
                            ParseDouble(parts[9]),  // T4_2
                            ParseDouble(parts[10]), // T4_3
                            ParseDouble(parts[11]), // T4_4
                            ParseDouble(parts[12]), // T4_5
                            ParseDouble(parts[13])  // T4_6

                        }
                    };
                    points.Add(dp);
                }
                catch (Exception ex)
                {
                    throw new FormatException($"Ошибка при разборе строки: {line}", ex);
                }
            }

            if (points.Count == 0)
                throw new InvalidDataException("Файл не содержит данных.");

            return points;
        }

        private static double ParseDouble(string s)
        {
            return double.Parse(s, System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
