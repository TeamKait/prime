using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace _38
{
    class Temperature
    {
        public int day;
        public int temp;

        public Temperature(int day, int temp)
        {
            this.day = day;
            this.temp = temp;
        }
    }
    internal class Program
    {
        static String _DIR = @"C:\Users\zadaa\Desktop\prime-main";
        static string filename(string name, string separator = @"\")
        {
            return _DIR + separator + name;
        }
        static void openSheetAndProcess(Excel.Workbook xlWorkbook, ref Collection<Temperature> temperatures, int sheet)
        {
            Console.WriteLine("Открываю лист номер " + sheet);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheet];
            xlWorksheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;


            Excel.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            int j = 3;
            Temperature currentTemp = new Temperature(0, 0);
            Console.WriteLine("Начинаю парсинг");
            for (int i = 1; i <= rowCount; i++)
            {

                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    try
                    {
                        string temp = xlRange.Cells[i, 1].Value2.ToString();
                        if (!double.TryParse(temp, out double _))
                        {
                            int date = Int32.Parse(temp.Split('.')[0]);
                            currentTemp.day = date;

                            temperatures.Add(currentTemp);
                            currentTemp = new Temperature(0, 0);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("Т");
                    }

                }

                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    try
                    {
                        int temp = (int)xlRange.Cells[i, j].Value2;

                        currentTemp.temp = temp;
                    }
                    catch
                    {
                        Console.WriteLine("Т");
                    }

                }
                Console.WriteLine("Парсинг строки " + i + " завершен");
            }

            Console.WriteLine("Завершил парсинг листа.");
        }
        static void ExcelExport(string file, dynamic[,] data)
        {
            object oMissing = System.Reflection.Missing.Value;

            Excel.Application excelApp = null;
            Excel.Range range = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            int worksheetCount = 0;

            try
            {

                excelApp = new Excel.Application();

                excelApp.DisplayAlerts = false;


                workbook = excelApp.Workbooks.Add();

                worksheetCount = workbook.Sheets.Count;

                worksheet = workbook.Sheets.Add();
                worksheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;


                if (data != null)
                {
                    for (int i = 0; i < data.GetLength(0); i++)
                    {
                        int rowNum = i + 1;

                        for (int j = 0; j < data.GetLength(1); j++)
                        {
                            int colNum = j + 1;
                            worksheet.Cells[rowNum, colNum] = data[i, j];
                        }
                    }
                }

                workbook.SaveAs(filename(file), System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

                Console.WriteLine("Сохранил файл " + file);

            }
            catch (Exception ex)
            {
                string errMsg = "Error (WriteToExcel) - " + ex.Message;
                Console.WriteLine(errMsg);
                System.Diagnostics.Debug.WriteLine(errMsg);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();

                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();

                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                }
            }


        }
        static void Main(string[] args)
        {
            Utils.info = new Utils.ProgramInfo(
              author: "Дерин Владислав",
              name: "38. Статистический анализ",
              description: "Работа с Excel",
              instruction: "");
            //Utils.PrintAuthors();
            Console.WriteLine("Открываю файл...");
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename("2019.xlsx"));
            Console.WriteLine("Файл открыт");
            Collection<Temperature> temperatures = new Collection<Temperature>();
            openSheetAndProcess(xlWorkbook, ref temperatures, 1);
            openSheetAndProcess(xlWorkbook, ref temperatures, 2);
            //openSheetAndProcess(xlWorkbook, ref temperatures, 3);
            //openSheetAndProcess(xlWorkbook, ref temperatures, 4);
            temperatures.Count.Output("Count");

            xlWorkbook.Close();
            xlApp.Quit();

            dynamic[,] data = new dynamic[32, temperatures.Count + 2];
            data[0, 0] = "Сравнение температур по датам четырех месяцев";
            data[0, 0] = "asdasdasd";


            for (int i = 1; i <= 31; i++)
            {
                data[i, 0] = i;

                foreach (Temperature temperature in temperatures)
                {
                    if (temperature.day == i)
                    {
                        int b = 1;
                        while (data[i, b] != null)
                        {
                            b++;
                        }
                        data[i, b] = temperature.temp;

                    }
                }
            }

            int maxVal = Int32.MinValue;
            for (int i = 1; i <= 31; i++)
            {
                int b = 1;
                int sum = 0;
                while (data[i, b] != null)
                {
                    sum+=data[i, b];
                    b++;
                }
                if(b > maxVal)
                {
                    maxVal = b;
                }
                data[i, maxVal+1] = Math.Round((double)sum/(double)(b-1), 2);
            }
            data[0, maxVal + 1] = "Среднее значение за день";

            Console.WriteLine("Введите имя файла: ");
            string file = Console.ReadLine();

            try
            {

                ExcelExport(file + ".xlsx", data);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            /*
            
            openSheetAndProcess(xlWorkbook, ref temperatures, 2);
            openSheetAndProcess(xlWorkbook, ref temperatures, 3);
            openSheetAndProcess(xlWorkbook, ref temperatures, 4);

            Console.WriteLine(temperatures.Count);
            Console.WriteLine("Введите имя файла: ");
            string file = Console.ReadLine();
            dynamic[,] data = new dynamic[32, temperatures.Count + 2];
            data[0, 0] = "Сравнение температур по датам четырех месяцев";
             */
        }
    }
}
