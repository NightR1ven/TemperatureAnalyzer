using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TemperatureAnalyzer.Models;

namespace TemperatureAnalyzer.Services
{
    /// <summary>
    /// Вычисление характеристик температурного поля по исходным точкам
    /// </summary>
    
    public static class ThermalCalculator
    {
        public static ThermalResult Calculate(List<DataPoint> points, double pH,
            string productNumber, double gasDensity, double stoichiometricRatio)
        {
            int n = points.Count;
            var result = new ThermalResult();

            result.ProductNumber = productNumber;
            result.GasDensity = gasDensity;
            result.StoichiometricRatio = stoichiometricRatio;

            double[] sum_t4cpi = new double[6];
            double[] max_t4 = new double[6];
            double[] min_t4 = new double[6];

            for (int i = 0; i < 6; i++)
            {
                max_t4[i] = double.MinValue;
                min_t4[i] = double.MaxValue;
                for (int j = 0; j < n; j++)
                {
                    double val = points[j].t4[i];
                    sum_t4cpi[i] += val;
                    if (val > max_t4[i]) max_t4[i] = val;
                    if (val < min_t4[i]) min_t4[i] = val;
                }
                result.t4cpi[i] = sum_t4cpi[i] / n;
                result.t4maxi[i] = max_t4[i];
                result.t4mini[i] = min_t4[i];
            }

            result.t4cp = result.t4cpi.Average();
            result.T4cp = result.t4cp + 273.15;

            for (int i = 0; i < 6; i++)
            {
                result.T4cpi[i] = result.t4cpi[i] + 273.15;
                result.T_4cpi[i] = result.T4cpi[i] / result.T4cp;
                result.dt4maxi[i] = result.t4maxi[i] - result.t4cpi[i];
                result.dt4mini[i] = result.t4mini[i] - result.t4cpi[i];

                // Для относительных подогревов используем tвх = 20 °C (комнатная)
                double tbbx = 20.0;
                double deltaT = result.t4cp - tbbx;
                if (Math.Abs(deltaT) > 1e-6)
                {
                    result.Ocpi[i] = (result.t4cpi[i] - tbbx) / deltaT;
                    result.Omaxi[i] = (result.t4maxi[i] - tbbx) / deltaT;
                }
                else
                {
                    result.Ocpi[i] = 0;
                    result.Omaxi[i] = 0;
                }
                result.dOi[i] = result.Omaxi[i] - result.Ocpi[i];
            }

            // Режимные параметры не измерены, оставляем 0
            result.tbbxcp = 0;
            result.Cbxcp = 0;
            result.lamcp = 0;
            result.alfacp = 0;
            result.Gbcp = 0;
            result.Ggcp = 0;
            result.Pbbxcp = 0;
            result.Pgcp = 0;

            return result;
        }
    }
}
