using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;

namespace ExcelFilterApp
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // Для новых версий EPPlus при необходимости настрой лицензию:
            ExcelPackage.License.SetNonCommercialPersonal("ASKO");
        }

        private void ProcessExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var ofd = new OpenFileDialog
                {
                    Filter = "Excel файлы|*.xlsx;*.xlsm;*.xls",
                    Title = "Выберите файл Excel"
                };

                if (ofd.ShowDialog() != true)
                    return;

                var fi = new FileInfo(ofd.FileName);

                using (var package = new ExcelPackage(fi))
                {
                    var wsSrc = package.Workbook.Worksheets.FirstOrDefault();
                    if (wsSrc == null || wsSrc.Dimension == null)
                    {
                        MessageBox.Show("Первый лист пустой или отсутствует.");
                        return;
                    }

                    // --- создаём/очищаем листы -K и -S ---
                    var wsK = package.Workbook.Worksheets["-K"] ?? package.Workbook.Worksheets.Add("-K");
                    wsK.Cells.Clear();

                    var wsS = package.Workbook.Worksheets["-S"] ?? package.Workbook.Worksheets.Add("-S");
                    wsS.Cells.Clear();

                    // --- копируем верхнюю строку (заголовок) до 5-го столбца ---
                    wsSrc.Cells[1, 1, 1, 5].Copy(wsK.Cells[1, 1, 1, 5]);
                    wsSrc.Cells[1, 1, 1, 5].Copy(wsS.Cells[1, 1, 1, 5]);

                    int startRow = 2;                 // пропускаем заголовок
                    int endRow = wsSrc.Dimension.End.Row;

                    int srcCol1 = 1;                 // исходный столбец 1 (типы/имена)
                    int srcCol5 = 5;                 // исходный столбец 5 (например, диаметр)

                    int dstColA = 1;                 // на целевых листах: в колонку 1 пишем srcCol1
                    int dstColB = 2;                 // в колонку 2 пишем srcCol5

                    int targetRowK = 2;              // первая строка данных на -K
                    int targetRowS = 2;              // первая строка данных на -S

                    // --- перенос строк с -K и -S ---
                    for (int row = startRow; row <= endRow; row++)
                    {
                        string text = wsSrc.Cells[row, srcCol1].Text;

                        if (string.IsNullOrWhiteSpace(text))
                            continue;

                        bool hasK = text.Contains("-K");
                        bool hasS = text.Contains("-S");

                        if (hasK)
                        {
                            wsK.Cells[targetRowK, dstColA].Value = wsSrc.Cells[row, srcCol1].Value;
                            wsK.Cells[targetRowK, dstColB].Value = wsSrc.Cells[row, srcCol5].Value;
                            targetRowK++;
                        }

                        if (hasS)
                        {
                            wsS.Cells[targetRowS, dstColA].Value = wsSrc.Cells[row, srcCol1].Value;
                            wsS.Cells[targetRowS, dstColB].Value = wsSrc.Cells[row, srcCol5].Value;
                            targetRowS++;
                        }
                    }

                    // --- сортировка по убыванию по первой цифре перед точкой в первом столбце ---
                    // функция: вытянуть число перед точкой, если есть, иначе -1 (чтобы строки без цифры ушли вниз)
                    int ExtractLeadingNumber(string s)
                    {
                        if (string.IsNullOrWhiteSpace(s))
                            return -1;

                        var parts = s.Split('.');
                        if (parts.Length > 1 && int.TryParse(parts[0], out int num))
                            return num;

                        return -1;
                    }

                    // сортировка листа -K
                    if (targetRowK > 2)
                    {
                        int lastDataRowK = targetRowK - 1;

                        var rowsK = new List<(int SortKey, string Col1, object Col2)>();
                        for (int r = 2; r <= lastDataRowK; r++)
                        {
                            string c1 = wsK.Cells[r, 1].Text;
                            object c2 = wsK.Cells[r, 2].Value;
                            int key = ExtractLeadingNumber(c1);
                            rowsK.Add((key, c1, c2));
                        }

                        var sortedK = rowsK
                            .OrderByDescending(x => x.SortKey) // 5,4,3,2,1,-1
                            .ThenBy(x => x.Col1)               // дополнительный порядок внутри одинаковых ключей
                            .ToList();

                        int writeRowK = 2;
                        foreach (var row in sortedK)
                        {
                            wsK.Cells[writeRowK, 1].Value = row.Col1;
                            wsK.Cells[writeRowK, 2].Value = row.Col2;
                            writeRowK++;
                        }
                    }

                    // сортировка листа -S
                    if (targetRowS > 2)
                    {
                        int lastDataRowS = targetRowS - 1;

                        var rowsS = new List<(int SortKey, string Col1, object Col2)>();
                        for (int r = 2; r <= lastDataRowS; r++)
                        {
                            string c1 = wsS.Cells[r, 1].Text;
                            object c2 = wsS.Cells[r, 2].Value;
                            int key = ExtractLeadingNumber(c1);
                            rowsS.Add((key, c1, c2));
                        }

                        var sortedS = rowsS
                            .OrderByDescending(x => x.SortKey)
                            .ThenBy(x => x.Col1)
                            .ToList();

                        int writeRowS = 2;
                        foreach (var row in sortedS)
                        {
                            wsS.Cells[writeRowS, 1].Value = row.Col1;
                            wsS.Cells[writeRowS, 2].Value = row.Col2;
                            writeRowS++;
                        }
                    }

                    // --- сохраняем изменения ---
                    package.Save();
                }

                MessageBox.Show("Готово: строки перенесены и отсортированы по убыванию (5,4,3,2,1, без цифр) на листах -K и -S.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}