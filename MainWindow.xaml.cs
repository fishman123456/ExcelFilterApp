using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;


namespace ExcelFilterApp
{
    public partial class MainWindow : Window
    {
        [Obsolete]
        public MainWindow()
        {
            InitializeComponent();
            // EPPlus требует указать контекст лицензии
            ExcelPackage.License.SetNonCommercialPersonal("ASKO");
        }

        private void ProcessExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 1. Диалог выбора Excel-файла
                var ofd = new OpenFileDialog
                {
                    Filter = "Excel файлы|*.xlsx;*.xlsm;*.xls",
                    Title = "Выберите файл Excel"
                };

                if (ofd.ShowDialog() != true)
                    return;

                var filePath = ofd.FileName;
                var fi = new FileInfo(filePath);

                // 2. Работа с книгой через EPPlus
                using (var package = new ExcelPackage(fi))
                {
                    var wsSrc = package.Workbook.Worksheets.FirstOrDefault();
                    if (wsSrc == null)
                    {
                        MessageBox.Show("В книге нет листов.");
                        return;
                    }

                    if (wsSrc.Dimension == null)
                    {
                        MessageBox.Show("Первый лист пустой.");
                        return;
                    }

                    // 3. Создаём или очищаем листы -K и -S
                    var wsK = package.Workbook.Worksheets["-K"];
                    if (wsK == null)
                        wsK = package.Workbook.Worksheets.Add("-K");
                    else
                        wsK.Cells.Clear();

                    var wsS = package.Workbook.Worksheets["-S"];
                    if (wsS == null)
                        wsS = package.Workbook.Worksheets.Add("-S");
                    else
                        wsS.Cells.Clear();

                    // 4. Копируем верхнюю строку первого листа как заголовок
                    // здесь предполагаем, что заголовок до 5-го столбца;
                    // при необходимости увеличь до нужного номера.
                    wsSrc.Cells[1, 1, 1, 5].Copy(wsK.Cells[1, 1, 1, 5]);
                    wsSrc.Cells[1, 1, 1, 5].Copy(wsS.Cells[1, 1, 1, 5]);

                    // 5. Настройки диапазона и столбцов
                    int startRow = 2;                      // пропускаем заголовок
                    int endRow = wsSrc.Dimension.End.Row;

                    int srcCol1 = 1;                      // исходный столбец 1
                    int srcCol5 = 5;                      // исходный столбец 5

                    // целевые столбцы: 1 и 2 (чтобы шло подряд)
                    int dstColA = 1;                      // значение из srcCol1
                    int dstColB = 2;                      // значение из srcCol5

                    int targetRowK = 2;                   // первая строка под данными на -K
                    int targetRowS = 2;                   // первая строка под данными на -S

                    // 6. Проход по строкам первого листа и фильтрация по -K / -S
                    for (int row = startRow; row <= endRow; row++)
                    {
                        string text = wsSrc.Cells[row, srcCol1].Text;

                        if (string.IsNullOrWhiteSpace(text))
                            continue;

                        bool hasK = text.Contains("-K");
                        bool hasS = text.Contains("-S");

                        // строки с "-K" → лист -K
                        if (hasK)
                        {
                            wsK.Cells[targetRowK, dstColA].Value =
                                wsSrc.Cells[row, srcCol1].Value;

                            wsK.Cells[targetRowK, dstColB].Value =
                                wsSrc.Cells[row, srcCol5].Value;

                            targetRowK++;
                        }

                        // строки с "-S" → лист -S
                        if (hasS)
                        {
                            wsS.Cells[targetRowS, dstColA].Value =
                                wsSrc.Cells[row, srcCol1].Value;

                            wsS.Cells[targetRowS, dstColB].Value =
                                wsSrc.Cells[row, srcCol5].Value;

                            targetRowS++;
                        }
                    }

                    // 7. Сохраняем изменения в тот же файл
                    package.Save();
                }

                MessageBox.Show("Готово: заголовок скопирован, строки с \"-K\" и \"-S\" перенесены.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}
