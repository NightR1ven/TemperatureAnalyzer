using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using TemperatureAnalyzer.Models;

namespace TemperatureAnalyzer.Services
{
    public static class ExcelExporter
    {
        public static void Export(ThermalResult result, List<DataPoint> allPoints, string filePath)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                // 1. Лист "Таблица" – статистика по поясам
                ExcelWorksheet wsResult = package.Workbook.Worksheets.Add("Таблица");
                FillResultSheet(wsResult, result);

                // 2. Лист "Эпюра" – график средних и максимальных значений по поясам
                ExcelWorksheet wsChart = package.Workbook.Worksheets.Add("Эпюра");
                CreateAverageMaxChart(wsChart, result);

                // 3. Лист "Исходные данные" – все точки (для справки)
                ExcelWorksheet wsData = package.Workbook.Worksheets.Add("Исходные данные");
                FillDataSheet(wsData, allPoints);

                // 4. Лист с полной эпюрой (все точки, все пояса) – один график
                CreateFullChartSheet(package, allPoints, result.ProductNumber);

                // 5. Лист с развёрткой на двух графиках (1-75 и 76-140) – два графика на одном листе
                CreateSplitChartSheet(package, allPoints, result.ProductNumber);

                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }
        }

        // ---------------------------------------------------------------
        // 1. Заполнение листа "Таблица"
        // ---------------------------------------------------------------
        private static void FillResultSheet(ExcelWorksheet ws, ThermalResult r)
        {
            int row = 1;
            ws.Cells[row, 1, row, 11].Merge = true;
            ws.Cells[row, 1].Value = "Характеристики температурного поля";
            ws.Cells[row, 1].Style.Font.Bold = true;
            ws.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            row += 2;

            string[] headers = { "Пояс", "t₄maxi,°C", "t₄mini,°C", "t₄cpi,°C", "Δt₄maxi,°C", "Δt₄mini,°C",
                                 "T₄cpi,K", "T̅₄cpi", "Θcpi", "Θmaxi", "ΔΘi" };
            for (int col = 0; col < headers.Length; col++)
                ws.Cells[row, col + 1].Value = headers[col];
            ws.Cells[row, 1, row, headers.Length].Style.Font.Bold = true;
            row++;

            for (int i = 0; i < 6; i++)
            {
                ws.Cells[row, 1].Value = i + 1;
                ws.Cells[row, 2].Value = r.t4maxi[i];
                ws.Cells[row, 3].Value = r.t4mini[i];
                ws.Cells[row, 4].Value = r.t4cpi[i];
                ws.Cells[row, 5].Value = r.dt4maxi[i];
                ws.Cells[row, 6].Value = r.dt4mini[i];
                ws.Cells[row, 7].Value = r.T4cpi[i];
                ws.Cells[row, 8].Value = r.T_4cpi[i];
                ws.Cells[row, 9].Value = r.Ocpi[i];
                ws.Cells[row, 10].Value = r.Omaxi[i];
                ws.Cells[row, 11].Value = r.dOi[i];
                row++;
            }

            row += 2;
            ws.Cells[row, 1].Value = "Средние значения режимных параметров:";
            ws.Cells[row, 1].Style.Font.Bold = true;
            row++;

            ws.Cells[row, 1].Value = "t₄cp,°C";
            ws.Cells[row, 2].Value = r.t4cp;
            ws.Cells[row, 3].Value = "T₄cp,K";
            ws.Cells[row, 4].Value = r.T4cp;
            row++;

            ws.Cells[row, 1].Value = "Номер изделия:";
            ws.Cells[row, 2].Value = r.ProductNumber;
            row++;
            ws.Cells[row, 1].Value = "Плотность газа, кг/м³:";
            ws.Cells[row, 2].Value = r.GasDensity;
            row++;
            ws.Cells[row, 1].Value = "Стехиометрический показатель:";
            ws.Cells[row, 2].Value = r.StoichiometricRatio;

            ws.Cells[ws.Dimension.Address].AutoFitColumns();
        }

        // ---------------------------------------------------------------
        // 2. Лист "Эпюра" – средние и максимумы по поясам
        // ---------------------------------------------------------------
        private static void CreateAverageMaxChart(ExcelWorksheet ws, ThermalResult r)
        {
            ws.Cells[1, 1].Value = "Пояс";
            ws.Cells[1, 2].Value = "t₄cpi,°C";
            ws.Cells[1, 3].Value = "t₄maxi,°C";
            for (int i = 0; i < 6; i++)
            {
                ws.Cells[i + 2, 1].Value = i + 1;
                ws.Cells[i + 2, 2].Value = r.t4cpi[i];
                ws.Cells[i + 2, 3].Value = r.t4maxi[i];
            }

            ExcelChart chart = ws.Drawings.AddChart("avgMaxChart", eChartType.LineMarkers);
            chart.SetSize(800, 400);
            chart.SetPosition(0, 0);
            chart.Title.Text = "Распределение температуры по поясам (средние и максимумы)";
            chart.Legend.Position = eLegendPosition.Bottom;

            var seriesAvg = chart.Series.Add(ws.Cells[2, 2, 7, 2], ws.Cells[2, 1, 7, 1]);
            seriesAvg.Header = "t₄cpi,°C";
            var seriesMax = chart.Series.Add(ws.Cells[2, 3, 7, 3], ws.Cells[2, 1, 7, 1]);
            seriesMax.Header = "t₄maxi,°C";
        }

        // ---------------------------------------------------------------
        // 3. Лист "Исходные данные"
        // ---------------------------------------------------------------
        private static void FillDataSheet(ExcelWorksheet ws, List<DataPoint> points)
        {
            int row = 1;
            ws.Cells[row, 1].Value = "№ точки";
            for (int i = 1; i <= 6; i++)
                ws.Cells[row, 1 + i].Value = $"t4_{i}";
            ws.Cells[row, 1, row, 7].Style.Font.Bold = true;
            row++;

            for (int i = 0; i < points.Count; i++)
            {
                ws.Cells[row, 1].Value = i + 1;
                for (int j = 0; j < 6; j++)
                    ws.Cells[row, 2 + j].Value = points[i].t4[j];
                row++;
            }
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
        }

        // ---------------------------------------------------------------
        // 4. Полная эпюра (все точки, один график) – числовая ось X, шаг 5
        // ---------------------------------------------------------------
        private static void CreateFullChartSheet(ExcelPackage package, List<DataPoint> points, string productNumber)
        {
            if (points == null || points.Count == 0) return;

            ExcelWorksheet dataSheet = GetOrCreateDataSheet(package, points);
            Color[] beltColors = GetBeltColors();

            ExcelWorksheet chartSheet = package.Workbook.Worksheets.Add("НК-16-18СТ (полная)");

            // Точечная диаграмма с линиями и маркерами
            ExcelChart chart = chartSheet.Drawings.AddChart("fullChart", eChartType.XYScatterLines);
            chart.SetSize(950, 500);
            chart.SetPosition(1, 0);
            chart.Title.Text = "Окружная неравномерность температурного поля (все точки)";
            chart.Legend.Position = eLegendPosition.Bottom;
            chart.XAxis.Title.Text = "Номер точки замера";
            chart.YAxis.Title.Text = "Температура, °C";

            // Настройка оси X: шаг 5, деления, минимальное значение = 0
            chart.XAxis.MajorUnit = 5;
            chart.XAxis.MinorUnit = 1;
            chart.XAxis.MajorTickMark = eAxisTickMark.Cross;
            chart.XAxis.MinValue = 0;

            for (int belt = 0; belt < 6; belt++)
            {
                var xRange = dataSheet.Cells[2, 1, points.Count + 1, 1];
                var yRange = dataSheet.Cells[2, belt + 2, points.Count + 1, belt + 2];
                var series = chart.Series.Add(yRange, xRange);
                series.Header = $"Т-{belt + 1}";
            }

            SetSeriesStyle(chart, beltColors);
            //AddTextAnnotations(chartSheet, productNumber);

            chartSheet.PrinterSettings.Orientation = eOrientation.Landscape;
            chartSheet.PrinterSettings.FitToPage = true;
            chartSheet.PrinterSettings.FitToWidth = 1;
            chartSheet.PrinterSettings.FitToHeight = 1;
        }

        // ---------------------------------------------------------------
        // 5. Развёртка на двух графиках (1-75 и 76-140) – графики не перекрываются
        // ---------------------------------------------------------------
        private static void CreateSplitChartSheet(ExcelPackage package, List<DataPoint> points, string productNumber)
        {
            if (points == null || points.Count == 0) return;

            ExcelWorksheet dataSheet = GetOrCreateDataSheet(package, points);
            Color[] beltColors = GetBeltColors();

            ExcelWorksheet splitSheet = package.Workbook.Worksheets.Add("Развёртка на 2 листах");
            splitSheet.PrinterSettings.Orientation = eOrientation.Landscape;

            int splitIndex = 75;
            if (points.Count < splitIndex) splitIndex = points.Count / 2;

            // ---- Первый график (1..splitIndex) ----
            ExcelChart chart1 = splitSheet.Drawings.AddChart("splitChart1", eChartType.XYScatterLines);
            chart1.SetSize(950, 280);
            chart1.SetPosition(1, 0);
            chart1.Title.Text = $"Точки 1 … {splitIndex}";
            chart1.Legend.Position = eLegendPosition.Bottom;
            chart1.XAxis.Title.Text = "Номер точки";
            chart1.YAxis.Title.Text = "Температура, °C";
            chart1.XAxis.MajorUnit = 5;
            chart1.XAxis.MinorUnit = 1;
            chart1.XAxis.MinValue = 1;

            for (int belt = 0; belt < 6; belt++)
            {
                var xRange = dataSheet.Cells[2, 1, splitIndex + 1, 1];
                var yRange = dataSheet.Cells[2, belt + 2, splitIndex + 1, belt + 2];
                var series = chart1.Series.Add(yRange, xRange);
                series.Header = $"Т-{belt + 1}";
            }
            SetSeriesStyle(chart1, beltColors);

            // ---- Второй график (splitIndex+1..конец) ----
            int start2 = splitIndex + 1;
            int end2 = points.Count;
            ExcelChart chart2 = splitSheet.Drawings.AddChart("splitChart2", eChartType.XYScatterLines);
            chart2.SetSize(950, 280);
            chart2.SetPosition(25, 0); // Смещение вниз на 25 строк – достаточно, чтобы не перекрывать первый
            chart2.Title.Text = $"Точки {start2} … {end2}";
            chart2.Legend.Position = eLegendPosition.Bottom;
            chart2.XAxis.Title.Text = "Номер точки";
            chart2.YAxis.Title.Text = "Температура, °C";
            chart2.XAxis.MajorUnit = 5;
            chart2.XAxis.MinorUnit = 1;
            chart2.XAxis.MinValue = start2;

            for (int belt = 0; belt < 6; belt++)
            {
                var xRange = dataSheet.Cells[start2 + 1, 1, end2 + 1, 1];
                var yRange = dataSheet.Cells[start2 + 1, belt + 2, end2 + 1, belt + 2];
                var series = chart2.Series.Add(yRange, xRange);
                series.Header = $"Т-{belt + 1}";
            }
            SetSeriesStyle(chart2, beltColors);

            //AddTextAnnotations(splitSheet, productNumber);

            splitSheet.PrinterSettings.FitToPage = true;
            splitSheet.PrinterSettings.FitToWidth = 1;
            splitSheet.PrinterSettings.FitToHeight = 1;
        }

        // ---------------------------------------------------------------
        // Вспомогательные методы
        // ---------------------------------------------------------------

        private static ExcelWorksheet GetOrCreateDataSheet(ExcelPackage package, List<DataPoint> points)
        {
            ExcelWorksheet dataSheet = package.Workbook.Worksheets["_Data"];
            if (dataSheet == null)
            {
                dataSheet = package.Workbook.Worksheets.Add("_Data");
                dataSheet.Hidden = eWorkSheetHidden.Hidden;

                dataSheet.Cells[1, 1].Value = "№ точки";
                for (int belt = 0; belt < 6; belt++)
                    dataSheet.Cells[1, belt + 2].Value = $"Т-{belt + 1}";

                for (int i = 0; i < points.Count; i++)
                {
                    dataSheet.Cells[i + 2, 1].Value = i + 1;
                    for (int belt = 0; belt < 6; belt++)
                        dataSheet.Cells[i + 2, belt + 2].Value = points[i].t4[belt];
                }
            }
            return dataSheet;
        }

        private static Color[] GetBeltColors()
        {
            return new Color[]
            {
                Color.FromArgb(0, 0, 255),      // Т-1 синий
                Color.FromArgb(255, 0, 0),      // Т-2 красный
                Color.FromArgb(0, 128, 0),      // Т-3 зелёный
                Color.FromArgb(128, 0, 128),    // Т-4 фиолетовый
                Color.FromArgb(0, 255, 255),    // Т-5 голубой
                Color.FromArgb(255, 165, 0)     // Т-6 оранжевый
            };
        }

        /// <summary>
        /// Применяет стили к сериям графика: цвет линии, тип маркера (круг), цвет маркера, размер маркера.
        /// </summary>
        private static void SetSeriesStyle(ExcelChart chart, Color[] beltColors)
        {
            int seriesCount = Math.Min(chart.Series.Count, beltColors.Length);
            if (seriesCount == 0) return;

            var chartXml = chart.ChartXml;
            var nsm = new System.Xml.XmlNamespaceManager(chartXml.NameTable);
            nsm.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            nsm.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var seriesNodes = chartXml.SelectNodes("//c:ser", nsm);
            for (int i = 0; i < seriesCount && i < seriesNodes.Count; i++)
            {
                var serElement = seriesNodes[i] as System.Xml.XmlElement;
                if (serElement == null) continue;

                var color = beltColors[i];

                // --- Настройка линии ---
                var spPrElement = serElement.SelectSingleNode("c:spPr", nsm) as System.Xml.XmlElement;
                if (spPrElement == null)
                {
                    spPrElement = chartXml.CreateElement("c", "spPr", nsm.LookupNamespace("c"));
                    serElement.AppendChild(spPrElement);
                }
                // Удаляем старый узел линии, если есть
                var oldLn = spPrElement.SelectSingleNode("a:ln", nsm);
                if (oldLn != null) spPrElement.RemoveChild(oldLn);

                var lnElement = chartXml.CreateElement("a", "ln", nsm.LookupNamespace("a"));
                lnElement.SetAttribute("w", "12700"); // толщина 1 pt
                var solidFill = chartXml.CreateElement("a", "solidFill", nsm.LookupNamespace("a"));
                var srgbClr = chartXml.CreateElement("a", "srgbClr", nsm.LookupNamespace("a"));
                srgbClr.SetAttribute("val", ColorToHex(color));
                solidFill.AppendChild(srgbClr);
                lnElement.AppendChild(solidFill);
                spPrElement.AppendChild(lnElement);

                // --- Настройка маркера ---
                var markerElement = serElement.SelectSingleNode("c:marker", nsm) as System.Xml.XmlElement;
                if (markerElement == null)
                {
                    markerElement = chartXml.CreateElement("c", "marker", nsm.LookupNamespace("c"));
                    serElement.AppendChild(markerElement);
                }
                else
                {
                    markerElement.RemoveAll();
                }
                // Тип маркера: круг
                var symbol = chartXml.CreateElement("c", "symbol", nsm.LookupNamespace("c"));
                symbol.SetAttribute("val", "circle");
                markerElement.AppendChild(symbol);
                // Размер маркера: 4
                var size = chartXml.CreateElement("c", "size", nsm.LookupNamespace("c"));
                size.SetAttribute("val", "4");
                markerElement.AppendChild(size);
                // Цвет маркера (заливка и контур)
                var markerSpPr = chartXml.CreateElement("c", "spPr", nsm.LookupNamespace("c"));
                var markerSolidFill = chartXml.CreateElement("a", "solidFill", nsm.LookupNamespace("a"));
                var markerSrgbClr = chartXml.CreateElement("a", "srgbClr", nsm.LookupNamespace("a"));
                markerSrgbClr.SetAttribute("val", ColorToHex(color));
                markerSolidFill.AppendChild(markerSrgbClr);
                markerSpPr.AppendChild(markerSolidFill);
                // Контур маркера
                var markerLn = chartXml.CreateElement("a", "ln", nsm.LookupNamespace("a"));
                markerLn.SetAttribute("w", "12700");
                var lnSolidFill = chartXml.CreateElement("a", "solidFill", nsm.LookupNamespace("a"));
                var lnSrgbClr = chartXml.CreateElement("a", "srgbClr", nsm.LookupNamespace("a"));
                lnSrgbClr.SetAttribute("val", ColorToHex(color));
                lnSolidFill.AppendChild(lnSrgbClr);
                markerLn.AppendChild(lnSolidFill);
                markerSpPr.AppendChild(markerLn);
                markerElement.AppendChild(markerSpPr);
            }
        }

        private static string ColorToHex(Color c)
        {
            return $"{c.R:X2}{c.G:X2}{c.B:X2}";
        }

        /// <summary>
        /// Добавляет аннотации в столбец K (11) с датой без ведущих нулей.
        /// </summary>
        //private static void AddTextAnnotations(ExcelWorksheet sheet, string productNumber)
        //{
        //    // Удаляем старые аннотации в ячейках, если они есть
        //    // (Просто очищаем диапазон, где мы их размещаем)
        //    var annotationRange = sheet.Cells[1, 11, 10, 13];
        //    annotationRange.Clear();

        //    // Записываем дату в столбец K (11)
        //    sheet.Cells[1, 11].Value = $"Дата: {DateTime.Now.ToString("d.M.yyyy")} г.";
        //    // Записываем номер изделия
        //    sheet.Cells[2, 11].Value = $"Номер изделия: {productNumber}";
        //    // Оставляем пустую строку для отступа
        //    sheet.Cells[3, 11].Value = "";
        //    // Записываем подписи
        //    sheet.Cells[4, 11].Value = "Обработал ______________";
        //    sheet.Cells[5, 11].Value = "Нач. стенда №3 ______________";
        //    sheet.Cells[6, 11].Value = "К/М цеха 420 ______________";

        //    // Настраиваем внешний вид ячеек с аннотациями (опционально)
        //    using (var range = sheet.Cells[1, 11, 6, 13])
        //    {
        //        range.Style.Font.Size = 10;
        //        range.Style.Font.Bold = false;
        //        range.Style.WrapText = true;
        //    }

        //    // Устанавливаем ширину столбцов, чтобы текст поместился
        //    sheet.Column(11).Width = 35;  // K
        //    sheet.Column(12).Width = 5;   // L
        //    sheet.Column(13).Width = 5;   // M
        //}
    }
}