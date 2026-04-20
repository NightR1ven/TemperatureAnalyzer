using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TemperatureAnalyzer.Models;

namespace TemperatureAnalyzer.Models
{
    /// <summary>
    /// Результаты расчёта характеристик температурного поля
    /// </summary>
    public class ThermalResult
    {
        // Массивы по 6 поясам
        public double[] t4maxi { get; set; } = new double[6];
        public double[] t4mini { get; set; } = new double[6];
        public double[] t4cpi { get; set; } = new double[6];
        public double[] dt4maxi { get; set; } = new double[6];
        public double[] dt4mini { get; set; } = new double[6];
        public double[] T4cpi { get; set; } = new double[6];
        public double[] T_4cpi { get; set; } = new double[6];
        public double[] Ocpi { get; set; } = new double[6];
        public double[] Omaxi { get; set; } = new double[6];
        public double[] dOi { get; set; } = new double[6];

        // Средние по всему сечению
        public double t4cp { get; set; }     // °C
        public double T4cp { get; set; }     // K

        // Средние режимные параметры (могут быть 0, если не измерены)
        public double tbbxcp { get; set; }   // °C
        public double Cbxcp { get; set; }    // м/с
        public double lamcp { get; set; }    // приведённая скорость
        public double alfacp { get; set; }   // коэффициент избытка воздуха
        public double Gbcp { get; set; }     // кг/с
        public double Ggcp { get; set; }     // кг/с
        public double Pbbxcp { get; set; }   // кгс/см²
        public double Pgcp { get; set; }     // кгс/см²

        // Параметры, вводимые пользователем
        public string ProductNumber { get; set; }
        public double GasDensity { get; set; }
        public double StoichiometricRatio { get; set; }

        // Для отображения в DataGrid
        public List<PoyasData> PoyasData
        {
            get
            {
                return Enumerable.Range(0, 6).Select(i => new PoyasData
                {
                    Index = i + 1,
                    t4maxi = t4maxi[i],
                    t4mini = t4mini[i],
                    t4cpi = t4cpi[i],
                    dt4maxi = dt4maxi[i],
                    dt4mini = dt4mini[i],
                    T4cpi = T4cpi[i],
                    T_4cpi = T_4cpi[i],
                    Ocpi = Ocpi[i],
                    Omaxi = Omaxi[i],
                    dOi = dOi[i]
                }).ToList();
            }
        }
    }

    public class PoyasData
    {
        public int Index { get; set; }
        public double t4maxi { get; set; }
        public double t4mini { get; set; }
        public double t4cpi { get; set; }
        public double dt4maxi { get; set; }
        public double dt4mini { get; set; }
        public double T4cpi { get; set; }
        public double T_4cpi { get; set; }
        public double Ocpi { get; set; }
        public double Omaxi { get; set; }
        public double dOi { get; set; }
    }
}