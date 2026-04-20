using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Printing;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using Microsoft.Win32;
using OxyPlot;
using OxyPlot.Series;
using TemperatureAnalyzer.Models;
using TemperatureAnalyzer.Services;
using TemperatureAnalyzer.Properties;
using System.Windows.Controls;

namespace TemperatureAnalyzer.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private ThermalResult _thermalResult;
        private int _selectedPointIndex;
        private List<Models.DataPoint> _allPoints;
        private double _pH;
        private string _statusText;
        private bool _isDataLoaded;
        private string _productNumber;
        private double _gasDensity;
        private double _stoichiometricRatio;

        public ThermalResult ThermalResult
        {
            get { return _thermalResult; }
            set { _thermalResult = value; OnPropertyChanged("ThermalResult"); }
        }

        public int SelectedPointIndex
        {
            get { return _selectedPointIndex; }
            set
            {
                if (value != _selectedPointIndex)
                {
                    _selectedPointIndex = value;
                    OnPropertyChanged("SelectedPointIndex");
                    Calculate();
                }
            }
        }

        public string StatusText
        {
            get { return _statusText; }
            set { _statusText = value; OnPropertyChanged("StatusText"); }
        }

        public bool IsDataLoaded
        {
            get { return _isDataLoaded; }
            set { _isDataLoaded = value; OnPropertyChanged("IsDataLoaded"); }
        }

        public string ProductNumber
        {
            get { return _productNumber; }
            set
            {
                if (_productNumber != value)
                {
                    _productNumber = value;
                    Settings.Default.ProductNumber = value;
                    Settings.Default.Save();
                    OnPropertyChanged("ProductNumber");
                }
            }
        }

        public double GasDensity
        {
            get { return _gasDensity; }
            set
            {
                if (Math.Abs(_gasDensity - value) > 1e-8)
                {
                    _gasDensity = value;
                    Settings.Default.GasDensity = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    Settings.Default.Save();
                    OnPropertyChanged("GasDensity");
                }
            }
        }

        public double StoichiometricRatio
        {
            get { return _stoichiometricRatio; }
            set
            {
                if (Math.Abs(_stoichiometricRatio - value) > 1e-8)
                {
                    _stoichiometricRatio = value;
                    Settings.Default.StoichiometricRatio = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    Settings.Default.Save();
                    OnPropertyChanged("StoichiometricRatio");
                }
            }
        }

        public PlotModel PlotModel { get; private set; }

        public ICommand LoadCommand { get; private set; }
        public ICommand CalculateCommand { get; private set; }
        public ICommand ExportCommand { get; private set; }
        public ICommand PrintCommand { get; private set; }

        public MainViewModel()
        {
            _productNumber = Settings.Default.ProductNumber;
            if (string.IsNullOrEmpty(_productNumber)) _productNumber = "БКС 257";

            double d;
            if (double.TryParse(Settings.Default.GasDensity, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out d))
                _gasDensity = d;
            else _gasDensity = 0.678;

            if (double.TryParse(Settings.Default.StoichiometricRatio, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out d))
                _stoichiometricRatio = d;
            else _stoichiometricRatio = 17.23577;

            LoadCommand = new RelayCommand(LoadData);
            CalculateCommand = new RelayCommand(Calculate, () => _allPoints != null);
            ExportCommand = new RelayCommand(ExportToExcel, () => ThermalResult != null);
            PrintCommand = new RelayCommand(PrintReport, () => ThermalResult != null);

            PlotModel = new PlotModel { Title = "Распределение температуры по поясам" };
            PlotModel.Axes.Add(new OxyPlot.Axes.LinearAxis { Position = OxyPlot.Axes.AxisPosition.Bottom, Title = "Номер пояса" });
            PlotModel.Axes.Add(new OxyPlot.Axes.LinearAxis { Position = OxyPlot.Axes.AxisPosition.Left, Title = "Температура, °C" });
            StatusText = "Загрузите файл ISH.txt";
        }

        private void LoadData()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*",
                Title = "Выберите файл ISH.txt"
            };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    _allPoints = DataReader.ReadData(dialog.FileName, out _pH);
                    SelectedPointIndex = _allPoints.Count;
                    StatusText = string.Format("Загружено {0} точек. Атм. давление: {1:F2} мм рт.ст.", _allPoints.Count, _pH);
                    IsDataLoaded = true;
                    Calculate();
                    CommandManager.InvalidateRequerySuggested();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Ошибка при загрузке файла: {0}", ex.Message), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusText = "Ошибка загрузки файла";
                    IsDataLoaded = false;
                }
            }
        }

        private void Calculate()
        {
            if (_allPoints == null || _allPoints.Count == 0) return;
            int n = SelectedPointIndex;
            if (n <= 0 || n > _allPoints.Count) n = _allPoints.Count;
            var points = _allPoints.Take(n).ToList();
            try
            {
                ThermalResult = ThermalCalculator.Calculate(points, _pH, ProductNumber, GasDensity, StoichiometricRatio);
                StatusText = string.Format("Расчёт выполнен по {0} точкам. t4cp = {1:F1} °C", n, ThermalResult.t4cp);
                UpdatePlot();
                CommandManager.InvalidateRequerySuggested();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Ошибка при расчёте: {0}", ex.Message), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText = "Ошибка расчёта";
            }
        }

        private void UpdatePlot()
        {
            PlotModel.Series.Clear();
            var seriesCpi = new LineSeries { Title = "t₄cpi (средняя)", MarkerType = MarkerType.Circle, MarkerSize = 4, StrokeThickness = 2 };
            var seriesMaxi = new LineSeries { Title = "t₄maxi (максимальная)", MarkerType = MarkerType.Triangle, MarkerSize = 4, StrokeThickness = 2, LineStyle = LineStyle.Dash };
            for (int i = 0; i < 6; i++)
            {
                seriesCpi.Points.Add(new OxyPlot.DataPoint(i + 1, ThermalResult.t4cpi[i]));
                seriesMaxi.Points.Add(new OxyPlot.DataPoint(i + 1, ThermalResult.t4maxi[i]));
            }
            PlotModel.Series.Add(seriesCpi);
            PlotModel.Series.Add(seriesMaxi);
            PlotModel.InvalidatePlot(true);
        }

        private void ExportToExcel()
        {
            if (ThermalResult == null) return;
            var dialog = new SaveFileDialog
            {
                Filter = "Excel файлы (*.xlsx)|*.xlsx",
                Title = "Сохранить протокол",
                FileName = string.Format("Протокол_Температурное_поле_{0:yyyyMMdd_HHmmss}.xlsx", DateTime.Now)
            };
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    ExcelExporter.Export(ThermalResult, _allPoints, dialog.FileName);
                    MessageBox.Show(string.Format("Протокол успешно сохранён в {0}", dialog.FileName), "Экспорт", MessageBoxButton.OK, MessageBoxImage.Information);
                    StatusText = string.Format("Протокол экспортирован в {0}", Path.GetFileName(dialog.FileName));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Ошибка при экспорте: {0}", ex.Message), "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void PrintReport()
        {
            var doc = new FlowDocument();
            var title = new Paragraph(new Run("Протокол испытаний камеры сгорания")) { FontSize = 18, FontWeight = System.Windows.FontWeights.Bold, TextAlignment = TextAlignment.Center };
            doc.Blocks.Add(title);
            var date = new Paragraph(new Run(string.Format("Дата расчёта: {0:dd.MM.yyyy HH:mm}", DateTime.Now))) { TextAlignment = TextAlignment.Right };
            doc.Blocks.Add(date);

            var table = new Table();
            table.BorderBrush = System.Windows.Media.Brushes.Black;
            table.BorderThickness = new Thickness(1);
            table.CellSpacing = 0;
            for (int i = 0; i < 7; i++) table.Columns.Add(new TableColumn { Width = new GridLength(i == 0 ? 60 : 80) });

            var headerRow = new TableRow();
            headerRow.Cells.Add(new TableCell(new Paragraph(new Run("Пояс"))) { FontWeight = System.Windows.FontWeights.Bold });
            for (int i = 1; i <= 6; i++) headerRow.Cells.Add(new TableCell(new Paragraph(new Run(i.ToString()))) { FontWeight = System.Windows.FontWeights.Bold });
            table.RowGroups.Add(new TableRowGroup());
            table.RowGroups[0].Rows.Add(headerRow);

            string[] paramNames = { "t₄maxi,°C", "t₄mini,°C", "t₄cpi,°C", "Δt₄maxi,°C", "Δt₄mini,°C",
                                    "T₄cpi,K", "T̅₄cpi", "Θcpi", "Θmaxi", "ΔΘi" };
            double[][] paramValues = {
                ThermalResult.t4maxi, ThermalResult.t4mini, ThermalResult.t4cpi,
                ThermalResult.dt4maxi, ThermalResult.dt4mini, ThermalResult.T4cpi,
                ThermalResult.T_4cpi, ThermalResult.Ocpi, ThermalResult.Omaxi, ThermalResult.dOi
            };
            for (int p = 0; p < paramNames.Length; p++)
            {
                var row = new TableRow();
                row.Cells.Add(new TableCell(new Paragraph(new Run(paramNames[p]))));
                for (int i = 0; i < 6; i++)
                {
                    string val = (p == 6) ? paramValues[p][i].ToString("F3") : paramValues[p][i].ToString("F2");
                    row.Cells.Add(new TableCell(new Paragraph(new Run(val))));
                }
                table.RowGroups[0].Rows.Add(row);
            }
            doc.Blocks.Add(table);

            doc.Blocks.Add(new Paragraph(new Run("Средние значения:")) { FontWeight = System.Windows.FontWeights.Bold, Margin = new Thickness(0, 10, 0, 0) });
            string avgText = string.Format("t₄cp = {0:F1} °C    T₄cp = {1:F1} K", ThermalResult.t4cp, ThermalResult.T4cp);
            doc.Blocks.Add(new Paragraph(new Run(avgText)));
            doc.Blocks.Add(new Paragraph(new Run(string.Format("Номер изделия: {0}", ThermalResult.ProductNumber))));
            doc.Blocks.Add(new Paragraph(new Run(string.Format("Плотность газа: {0:F3} кг/м³", ThermalResult.GasDensity))));
            doc.Blocks.Add(new Paragraph(new Run(string.Format("Стехиометрический показатель: {0:F5}", ThermalResult.StoichiometricRatio))));

            var printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                var paginator = ((IDocumentPaginatorSource)doc).DocumentPaginator;
                printDialog.PrintDocument(paginator, "Протокол температурного поля");
                StatusText = "Отправлено на печать";
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;
        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            if (execute == null) throw new ArgumentNullException("execute");
            _execute = execute;
            _canExecute = canExecute;
        }
        public bool CanExecute(object parameter) { return _canExecute == null ? true : _canExecute(); }
        public void Execute(object parameter) { _execute(); }
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}