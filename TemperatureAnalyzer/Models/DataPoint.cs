using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TemperatureAnalyzer.Models;

namespace TemperatureAnalyzer.Models
{
    /// <summary>
    /// Исходные данные для одной точки измерения (ISH.txt)
    /// </summary>
    public class DataPoint
    {
        // Поля 0-8 не используются в расчёте, но хранятся для полноты
        public double Field0 { get; set; }
        public double Field1 { get; set; }
        public double Field2 { get; set; }
        public double Field3 { get; set; }
        public double Field4 { get; set; }
        public double Field5 { get; set; }
        public double Field6 { get; set; }
        public double Field7 { get; set; }
        public double Field8 { get; set; }
        // Температуры по поясам (6 значений)
        public double[] t4 { get; set; } = new double[6];
    }
}