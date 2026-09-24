using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace mamaRabota.Classes
{
    internal class ExcelClass
    {
        Excel.Application app;
        private string File { get; set; }
        private string Month { get; set; }
        private double FullSum { get; set; }
        private int FullDays { get; set; }
        private double FullLitrs { get; set; }
        private string Season { get; set; }
        private int NuberList { get; set; }
        private int NuberList2 { get; set; }
        private double Target { get; set; }
        private List<System.Windows.Controls.CheckBox> AllChekBoxes { get; set; }
        private Dictionary<string, string> DaysLitrs { get; set; }

        private const int TEMPLATE_INDEX = 2;

        public ExcelClass(ExcelData data)
        {
            File = data.File;
            FullSum = data.FullSum;
            FullLitrs = data.FullLitrs;
            FullDays = data.FullDays;
            AllChekBoxes = data.AllChekBoxes;
            Month = data.Month;
            NuberList = data.NuberList;
            NuberList2 = data.NuberList2;
            DaysLitrs = data.DaysLitrs;
            Season = data.Season;
            Target= data.Target;
        }

        public void OpenFile()
        {
            Excel.Application app = null;
            Excel.Workbook workbook = null;
            try
            {
                app = new Excel.Application();
                workbook = app.Workbooks.Open(File);
                app.Visible = true;
                AdjustSheetsCount(workbook);
                AddMainData(workbook);
                SetDateDown(workbook);
                SetDateUp(workbook);
                SetNumberList(workbook);
                SetLitrsInOneList(workbook);
                app.CalculateFull();
                BalanceRemains(workbook);
                MessageBox.Show("✅ Файл успешно заполнен!", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии файла");
            }
        }
        private void AdjustSheetsCount(Workbook workbook)
        {
            int neededDays = AllChekBoxes.Count;
            int currentDaySheets = workbook.Worksheets.Count - 1;

            if (neededDays > currentDaySheets)
            {
                for (int i = 0; i < currentDaySheets; i++)
                {
                    Excel.Worksheet sheet = workbook.Worksheets[i + 2];
                    sheet.Name = $"{AllChekBoxes[i].Content}.{Month}";
                }
                for (int i = currentDaySheets; i < neededDays; i++)
                {
                    string previousSheetName = $"{AllChekBoxes[i - 1].Content}.{Month}";
                    string newSheetName = $"{AllChekBoxes[i].Content}.{Month}";

                    Excel.Worksheet lastSheet = workbook.Worksheets[workbook.Worksheets.Count];
                    lastSheet.Copy(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]);

                    Excel.Worksheet newSheet = workbook.Worksheets[workbook.Worksheets.Count];
                    UpdateFormulasInSheet(newSheet, previousSheetName);

                    newSheet.Name = newSheetName;
                }
            }
            else if (neededDays < currentDaySheets)
            {
                int toRemove = currentDaySheets - neededDays;
                for (int i = 0; i < toRemove; i++)
                {
                    workbook.Worksheets[workbook.Worksheets.Count].Delete();
                }

                for (int i = 0; i < neededDays; i++)
                {
                    workbook.Worksheets[i + 2].Name = $"{AllChekBoxes[i].Content}.{Month}";
                }
            }
            else
            {
                for (int i = 0; i < neededDays; i++)
                {
                    workbook.Worksheets[i + 2].Name = $"{AllChekBoxes[i].Content}.{Month}";
                }
            }
        }


        private void UpdateFormulasInSheet(Excel.Worksheet sheet, string correctPreviousName)
        {
            // Проходим по всем использованным ячейкам
            Excel.Range usedRange = sheet.UsedRange;

            foreach (Excel.Range cell in usedRange)
            {
                try
                {
                    if (cell.HasFormula)
                    {
                        string formula = cell.Formula.ToString();
                        string updatedFormula = ReplaceSheetReference(formula, correctPreviousName);
                        if (updatedFormula != formula)
                        {
                            cell.Formula = updatedFormula;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Ошибка чтения формулы");
                }
            }
        }

        /// <summary>
        /// Заменяет все ссылки на листы в формуле на correctName
        /// </summary>
        private string ReplaceSheetReference(string formula, string correctName)
        {
            var result = new System.Text.StringBuilder();
            int i = 0;

            while (i < formula.Length)
            {
                if (i + 1 < formula.Length && formula[i] == '=' && formula[i + 1] == '\'')
                {
                    int endQuote = formula.IndexOf("'!", i + 2);
                    if (endQuote != -1)
                    {
                        result.Append("='");
                        result.Append(correctName);
                        result.Append("'!");
                        i = endQuote + 2;
                        continue;
                    }
                }

                if (formula[i] == '=')
                {
                    int exclamation = formula.IndexOf('!', i + 1);
                    int space = formula.IndexOf(' ', i + 1);

                    if (exclamation != -1 && (space == -1 || exclamation < space))
                    {
                        string sheetRef = formula.Substring(i + 1, exclamation - i - 1);
                        if (sheetRef.Contains("."))
                        {
                            result.Append("=");
                            result.Append(correctName);
                            result.Append("!");
                            i = exclamation + 1;
                            continue;
                        }
                    }
                }

                result.Append(formula[i]);
                i++;
            }

            return result.ToString();
        }

        private void AddMainData(Workbook workbook)
        {
            Excel.Worksheet worksheet = workbook.Worksheets[TEMPLATE_INDEX];
            worksheet.Range["G2"].Value = FullDays;
            worksheet.Range["G3:H3"].Value = FullSum;
            worksheet.Range["G4"].Value = FullLitrs;
        }

        private void SetDateDown(Workbook workbook)
        {
            var year = DateTime.Now.Year;
            for (int i = 0; i < AllChekBoxes.Count; i++)
            {
                Excel.Worksheet sheet = workbook.Worksheets[i + TEMPLATE_INDEX];
                sheet.Range["B65:E65"].Value = $"{AllChekBoxes[i].Content}.{Month}.{year}";
            }
        }

        private void SetDateUp(Workbook workbook)
        {
            var year = DateTime.Now.Year;
            for (int i = 0; i < AllChekBoxes.Count; i++)
            {
                Excel.Worksheet sheet = workbook.Worksheets[i + TEMPLATE_INDEX];
                sheet.Range["A7:Q7"].Value = $"за период с {AllChekBoxes[i].Content} {GetMonthNameGenitive(Month)} {year}г. По {AllChekBoxes[i].Content} {GetMonthNameGenitive(Month)} {year}г.";
            }
        }
        private void SetNumberList(Workbook workbook)
        {
            for (int i = 0; i < AllChekBoxes.Count; i++)
            {
                Excel.Worksheet sheet = workbook.Worksheets[i + TEMPLATE_INDEX];
                sheet.Range["W5:Z5"].Value = $"{NuberList + i}-{NuberList2}";
            }
        }

        private string GetMonthNameGenitive(string monthNumber)
        {
            var months = new Dictionary<string, string>
            {
                {"01", "января"}, {"02", "февраля"}, {"03", "марта"},
                {"04", "апреля"}, {"05", "мая"}, {"06", "июня"},
                {"07", "июля"}, {"08", "августа"}, {"09", "сентября"},
                {"10", "октября"}, {"11", "ноября"}, {"12", "декабря"}
            };

            if (months.TryGetValue(monthNumber, out string monthName))
            {
                return monthName;
            }
            return "";
        }

        private void SetLitrsInOneList(Workbook workbook)
        {
            try
            {
                // Очищаем и заполняем сезон на всех листах дней
                for (int i = 0; i < AllChekBoxes.Count; i++)
                {
                    Excel.Worksheet sheet = workbook.Worksheets[i + TEMPLATE_INDEX];
                    sheet.Range["Y42"].Value = "";
                    sheet.Range["Y51"].Value = Season;
                }

                // Заполняем литры для конкретных дней
                foreach (var dayLitr in DaysLitrs)
                {
                    string targetName = $"{dayLitr.Key}.{Month}";
                    for (int i = 0; i < AllChekBoxes.Count; i++)
                    {
                        Excel.Worksheet sheet = workbook.Worksheets[i + TEMPLATE_INDEX];
                        if (sheet.Name == targetName)
                        {
                            sheet.Range["Y42"].Value = dayLitr.Value;
                            break;
                        }
                    }
                }
            }
            catch
            {
                MessageBox.Show("Ошибка заполнения дней заправки");
            }
        }

        public void BalanceRemains(Workbook workbook)
        {
            try
            {
                List<Excel.Worksheet> daySheets = new List<Excel.Worksheet>();
                for (int i = 0; i < AllChekBoxes.Count; i++)
                {
                    daySheets.Add(workbook.Worksheets[i + TEMPLATE_INDEX]);
                }

                if (daySheets.Count == 0) return;

                Excel.Worksheet lastSheet = daySheets.Last();
                double TARGET = Target;

                // 1. Пересчитываем, чтобы получить актуальные данные
                workbook.Application.CalculateFull();
                double currentRemains = Convert.ToDouble(lastSheet.Range["Y47"].Value ?? 0);
                double difference = currentRemains - TARGET;

  
                if (Math.Abs(difference) < 0.01) return;

                // 2. Вычисляем точный шаг корректировки
                double stepPerSheet = difference / daySheets.Count;


                string sign = stepPerSheet >= 0 ? "+" : "-";
                string absStepString = Math.Abs(stepPerSheet).ToString("F4", System.Globalization.CultureInfo.InvariantCulture);

                // 3. Применяем корректировку к КАЖДОМУ листу, сохраняя формулу
                foreach (var sheet in daySheets)
                {
                    var cell = sheet.Range["Y52"];
                    string baseFormula = "";

                    // А. Получаем текущую базу формулы
                    if (cell.HasFormula)
                    {
                        baseFormula = cell.Formula.ToString();
                        if (baseFormula.StartsWith("="))
                            baseFormula = baseFormula.Substring(1);
                    }
                    else
                    {
           
                        baseFormula = Convert.ToDouble(cell.Value ?? 0).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }

                    // Б. Формируем новую формулу: =(База) + Шаг
                    string newFormula = $"=({baseFormula}){sign}{absStepString}";
                    double currentVal = Convert.ToDouble(cell.Value ?? 0);
                    double projectedVal = currentVal + stepPerSheet;
                                                                    
                    double logicalStep = (stepPerSheet >= 0) ? Math.Abs(stepPerSheet) : -Math.Abs(stepPerSheet);
                    projectedVal = currentVal + logicalStep;

                    if (projectedVal < 0.1)
                    {
                        cell.Value = 0.1;
                    }
                    else
                    {
                        cell.Formula = newFormula;
                    }
                }

                // 4. Финальный пересчет
                workbook.Application.CalculateFull();

            
                double finalRemains = Convert.ToDouble(lastSheet.Range["Y47"].Value ?? 0);
                double finalDiff = finalRemains - TARGET;

                if (Math.Abs(finalDiff) > 0.01)
                {
                    // Корректируем последний лист на оставшуюся копейку
                    var lastCell = lastSheet.Range["Y52"];
                    string lastBase = "";
                    if (lastCell.HasFormula)
                    {
                        lastBase = lastCell.Formula.ToString().TrimStart('=');
                    }
                    else
                    {
                        lastBase = Convert.ToDouble(lastCell.Value ?? 0).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }

                    string lastSign = finalDiff >= 0 ? "+" : "-";
                    string lastStep = Math.Abs(finalDiff).ToString("F4", System.Globalization.CultureInfo.InvariantCulture);

                    lastCell.Formula = $"=({lastBase}){lastSign}{lastStep}";
                    workbook.Application.CalculateFull();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка балансировки остатка");
            }
        }


    }

    public class ExcelData
    {
        public string File { get; set; }
        public string Month { get; set; }
        public int FullDays { get; set; }
        public double FullSum { get; set; }
        public double FullLitrs { get; set; }
        public double Target { get; set; }
        public int NuberList { get; set; }
        public int NuberList2 { get; set; }
        public string Season { get; set; }
        public List<System.Windows.Controls.CheckBox> AllChekBoxes { get; set; }
        public Dictionary<string, string> DaysLitrs { get; set; }
    }
}