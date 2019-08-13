using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace ExcelCreate
{
    class Data
    {
        // Массив имен и сумм компаний
        public string[,] Companies = new string[,] {
            { "Телеком.ру", "1520.00" },
            { "ИП Иванов", "1020.50" },
            { "CARGO - BLG ТЭК", "1240.12" },
            { "АВАНГАРД - Техно", "1380.14" },
            { "НордСтройАвто", "1100.15" },
            { "АвтоГОСТ","2334.33"},
            { "Аванстрой","3417.44"},
            { "Норматех","2726.64"},
            { "Азия - Авто","1640.43"},
            { "АВАНТАЖ","3929.12"},
            { "НОТЗ","2433.64"},
            { "ИСТ","1834.97"},
            { "АВБ - ГРУПП Сервис","4043.34"},
            { "НПК Пермский механический завод","2508.19"},
            { "Склад карданных валов","1535.13"},
            { "АВБ - СЕРВИС","2552.76"},
            { "UkraineНПО Диапазон","4425.34"},
            { "СпецАвтоТранс","3986.31"},
            { "АВЕНЮ"," 4365.76"},
            { "НПО ПУЛЬС","3234.45"},
            { "ЮГТРАНССЕРВИС","2917.23"},
            { "АВЕРС ","1306.65"},
            { "НПО","3052.23"},
            { "СтройМехТранс","2095.54"},
            { "1000 Инструментов","3882.34"},
            { "АВЕРС ВН","2309.76"},
            { "НСК","2330.09"}};


        // Метод получения Excel документа
        public void GetVATExcel(string date, string VAT, string savePath)
        {

            Excel.Application ex = new Excel.Application();

            ex.SheetsInNewWorkbook = 1;

            // Отобразить документ?
            ex.Visible = false;

            Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);

            ex.DisplayAlerts = false;

            Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

            sheet.Name = "SumCompanies";

            // Для подсчета суммы НДС
            double VATForCalculation;
            string[] textVAT = new string[2];
            double[] sum = new double[Companies.GetLength(0)];
            CultureInfo ciEnUs = new CultureInfo("en-us");
            if (VAT == "20" || VAT == "18" || VAT == "16.67" || VAT == "10")
                VATForCalculation = double.Parse(VAT, ciEnUs) / 100;
            else
            {
                for (int i = 0; i < 2; i++)
                    textVAT = VAT.Split('/');
                VATForCalculation = Convert.ToDouble(textVAT[0]) / (Convert.ToDouble(textVAT[1]));
            }
            for (int i = 0; i < Companies.GetLength(0); i++)
            {
                sum[i] = double.Parse(Companies[i, 1], ciEnUs);
            }

            //Заполнение ячеек
            for (int i = 2; i < Companies.GetLength(0) + 2; i++)
            {
                sheet.Cells[i, 1] = date;
                sheet.Cells[i, 2] = String.Format(Companies[i - 2, 0]);
                sheet.Cells[i, 3] = ((int)sum[i - 2]);
                sheet.Cells[i, 4] = (VAT);
                sheet.Cells[i, 5] = (int)((sum[i - 2] * (VATForCalculation)));
            }

            //Заголовки
            sheet.Cells[1, 1] = "Дата";
            sheet.Cells[1, 2] = "Наименование";
            sheet.Cells[1, 3] = "Сумма,р";
            sheet.Cells[1, 4] = "НДС,%";
            sheet.Cells[1, 5] = "Сумма НДС,р";

            //Результат
            sheet.Cells[30, 2] = "ИТОГО";
            sheet.Cells[30, 3].Formulalocal = "=СУММ(C1:C28)";
            sheet.Cells[30, 4] = VAT;
            sheet.Cells[30, 5].Formulalocal = "=СУММ(E1:E28)";

            //Захват диапазон ячеек
            Excel.Range rangeDate = sheet.get_Range("A1", "A28");
            Excel.Range rangeName = sheet.get_Range("B1", "B28");
            Excel.Range rangeSum = sheet.get_Range("C1", "C28");
            Excel.Range rangeНДС = sheet.get_Range("D1", "D28");
            Excel.Range rangeSumНДС = sheet.get_Range("E1", "E28");
            sheet.get_Range("A1", "E1").Font.Bold = 1;
            sheet.get_Range("B30", "E30").Font.Bold = 1;

            //Размер шрифта для диапазона
            rangeDate.EntireColumn.AutoFit();
            rangeName.EntireColumn.AutoFit();
            rangeSum.EntireColumn.AutoFit();
            rangeНДС.EntireColumn.AutoFit();
            rangeSumНДС.EntireColumn.AutoFit();

            //Сохранение
            if (File.Exists(savePath) == false)
            {
                ex.Application.ActiveWorkbook.SaveAs(savePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlShared);
            }
            else
            {
                var result = MessageBox.Show("Файл уже существует. Перезаписать его?", "Ошибка", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                    ex.Application.ActiveWorkbook.SaveAs(savePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlShared);
                //Закрытие Excel
                ex.Application.Quit();
            }

        }
    }
}
