using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn1
{
    public static class PriceTag
    {        
        
        public enum PriceTagSize
        {
            Small = 0,
            Big = 1
        }


        /// <summary>
        /// Создает новый лсит с ценниками.
        /// </summary>
        /// <param name="tagSize"> Тип ценников. </param>
        public static void CreatePriceTagSheet(PriceTagSize tagSize)
        {
#if DEBUG
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
#endif
            int colCount = 3;
            double[] colsWidth = new double[1];
            double[] rowsHeight = new double[1];
            int[] fontSize = new int[1];
            FontOptions[] fontOptions = new FontOptions[1];

            // Установка параметров ценника в зависимости от типа
            switch (tagSize)
            {
                case PriceTagSize.Small:
                    colCount = 3;
                    colsWidth = new double[] {
                        14.29, 2.86, 10.14, 2.57,
                        14.29, 2.86, 10.14, 2.57,
                        14.29, 2.86, 10.14, 2.57 };

                    rowsHeight = new double[] { 43.50, 18.00, 28.50, 24.00 };
                    fontSize = new int[] { 12, 18, 14, 24, 20, 18, 20, 16, 8 };

                    fontOptions = new FontOptions[] {
                        new FontOptions(size: 22, length: 0, wrap: false),
                        new FontOptions(size: 20, length: 16, wrap: false),
                        new FontOptions(size: 18, length: 19, wrap: true),
                        new FontOptions(size: 16, length: 34, wrap: true),
                        new FontOptions(size: 14, length: 44, wrap: true),
                        new FontOptions(size: 12, length: 56, wrap: true) };
                    break;

                case PriceTagSize.Big:
                    colCount = 2;
                    colsWidth = new double[] {
                        23.14, 6.86, 19.43, 6.86,
                        23.14, 6.86, 19.43, 6.86 };

                    rowsHeight = new double[] { 57.75, 29.25, 42.75, 28.50 };
                    fontSize = new int[] { 14, 24, 22, 36, 28, 28, 28, 26, 11 };

                    fontOptions = new FontOptions[] {
                        new FontOptions(size: 36, length: 0, wrap: false),
                        new FontOptions(size: 28, length: 16, wrap: false),
                        new FontOptions(size: 24, length: 22, wrap: true),
                        new FontOptions(size: 22, length: 44, wrap: true),
                        new FontOptions(size: 20, length: 52, wrap: true),
                        new FontOptions(size: 18, length: 60, wrap: true),
                        new FontOptions(size: 16, length: 70, wrap: true),
                        new FontOptions(size: 14, length: 80, wrap: true) };
                    break;
            }

            var app = Globals.ThisAddIn.Application;

            app.ScreenUpdating = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;

            var firstSheet = app.ActiveSheet as Excel.Worksheet;
            object[,] values = firstSheet.Range["A1:G100"].Value2;

            Excel.Worksheet newWorksheet = (Excel.Worksheet)app.Worksheets.Add();
            // установить ширину столбиков на новом листе 
            newWorksheet.Range[$"A1:{(char)(colCount * 4 + 64)}1"].ColumnWidth = colsWidth;            

            // Создание ценников на новом листе
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (values[i, 1] != null)
                {
                    //int idx = cell.Row;
                    var targetCell = newWorksheet.Range[GetStartCell(i, colCount)];

                    if (i % colCount == 1)
                    {
                        targetCell.RowHeight = rowsHeight[0];
                        targetCell.Offset[1, 0].RowHeight = rowsHeight[1];
                        targetCell.Offset[2, 0].RowHeight = rowsHeight[2];
                        targetCell.Offset[3, 0].RowHeight = rowsHeight[3];
                    }
                    // массив данных для ценника: 0 - название, 1 - артикул, 2 - цена опт, 3 - цена розница
                    string[] info = new string[4];
                    info[0] = Convert.ToString(values[i, 4]);
                    info[1] = Convert.ToString(values[i, 3]);
                    info[2] = Convert.ToString(values[i, 6]);
                    info[3] = Convert.ToString(values[i, 7]);

                    fontSize[0] = AdjustFontSize(info[0].Length, fontOptions, out bool wrapNameCell);

                    CreatePriceTag(targetCell, info, fontSize, wrapNameCell);
                }
                else
                {
                    break;
                }
            }

            
            if (tagSize == PriceTagSize.Small)
            {
                // настроить поля
                newWorksheet.PageSetup.HeaderMargin = 0;
                newWorksheet.PageSetup.FooterMargin = 0;
                double margin = newWorksheet.Application.CentimetersToPoints(0.5);
                newWorksheet.PageSetup.TopMargin = margin;
                newWorksheet.PageSetup.BottomMargin = margin;
                newWorksheet.PageSetup.LeftMargin = margin;
                newWorksheet.PageSetup.RightMargin = margin;
            }

            app.ScreenUpdating = true;
            app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

#if DEBUG
            stopwatch.Stop();
            newWorksheet.Range["M1"].Value2 = stopwatch.ElapsedMilliseconds;
#endif
        }

        
        /// <summary>
        /// Создает ценник.
        /// </summary>
        /// <param name="firstCell"> Левая верхняя ячейка ценника. </param>
        /// <param name="info"> Массив входных данных. </param>
        /// <param name="fontSize"> Размеры шрифтов. </param>
        /// <param name="wrapNameCell"> Перенос текста в ячейке наименования. </param>
        private static void CreatePriceTag(Excel.Range firstCell, string[] info, int[] fontSize, bool wrapNameCell)
        {
            var workSheet = firstCell.Worksheet;

            // граница всего ценника
            var range = workSheet.Range[firstCell, firstCell.Offset[3, 3]];
            range.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
            // общее выравнивание
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // линия между ценами
            range = workSheet.Range[firstCell.Offset[1, 1], firstCell.Offset[2, 1]];
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThick;


            // объединение ячеек
            workSheet.Range[firstCell.Offset[1, 0], firstCell.Offset[1, 1]].Merge();
            workSheet.Range[firstCell.Offset[1, 2], firstCell.Offset[1, 3]].Merge();

            workSheet.Range[firstCell.Offset[3, 0], firstCell.Offset[3, 1]].Merge();
            workSheet.Range[firstCell.Offset[3, 2], firstCell.Offset[3, 3]].Merge();

            // строка 2
            var curCell = firstCell.Offset[1, 0];
            curCell.Value2 = "Цена опт";
            curCell.Font.Size = fontSize[1];
            curCell.Font.Bold = true;

            curCell = firstCell.Offset[1, 2];
            curCell.Value2 = "Цена розн.";
            curCell.Font.Size = fontSize[2];

            // цена опт
            curCell = firstCell.Offset[2, 0];
            curCell.Value2 = double.Parse(info[2]);
            curCell.NumberFormat = "0.00";
            curCell.Font.Size = fontSize[3];
            curCell.Font.Bold = true;

            // цена розница
            curCell = firstCell.Offset[2, 2];
            curCell.Value2 = double.Parse(info[3]);
            curCell.NumberFormat = "0.00";
            curCell.Font.Size = fontSize[5];


            // значок рубля
            firstCell.Offset[2, 1].Value2 = "₽";
            firstCell.Offset[2, 1].Font.Size = fontSize[4];

            firstCell.Offset[2, 3].Value2 = "₽";
            firstCell.Offset[2, 3].Font.Size = fontSize[6];

            // артикул
            curCell = firstCell.Offset[3, 0];
            curCell.Value2 = info[1];
            curCell.Font.Size = fontSize[7];
            curCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            // флп
            firstCell.Offset[3, 2].Value2 = "ФЛП Никулин Ю.В.";
            firstCell.Offset[3, 2].Font.Size = fontSize[8];

            

            // объединять первую строчку в конце
            workSheet.Range[firstCell, firstCell.Offset[0, 3]].Merge();

            // название товара            
            firstCell.Value2 = info[0];
            firstCell.Font.Size = fontSize[0];
            firstCell.WrapText = wrapNameCell;
            firstCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;            
        }


        /// <summary>
        /// Возвращает подходящий размер шрифта.
        /// </summary>
        /// <param name="textLength"> Длина текста. </param>
        /// <param name="fontOptions"> Набор соответствий длины текста и размера шрифта. </param>
        /// <param name="wrapNameCell"> Перенос текста в ячейке. </param>
        /// <returns></returns>
        private static int AdjustFontSize(int textLength, FontOptions[] fontOptions, out bool wrapNameCell)
        {
            int fontSize = fontOptions[0].Size;
            bool boolResult = fontOptions[0].Wrap;

            for (int i = 1; i < fontOptions.Length; i++)
            {
                if (textLength > fontOptions[i].Length)
                {
                    fontSize = fontOptions[i].Size;
                    boolResult = fontOptions[i].Wrap;
                }
            }

            wrapNameCell = boolResult;
            return fontSize;
        }

        
        /// <summary>
        /// Возвращает имя начальной ячейки
        /// </summary>
        /// <param name="idx"> Порядковый номер, начиная с 1. </param>
        /// <param name="colCount"> Количество столбцов в выходной таблице. </param>
        /// <returns></returns>
        private static string GetStartCell(int idx, int colCount)
        {
            // A = 65
            const int width = 4;
            const int height = 4;

            int row = (idx - 1) / colCount;
            int col = (idx - 1) % colCount;

            char excelCol = (char)(col * width + 65);

            return $"{excelCol}{row * height + 1}";
        }


        /// <summary>
        /// Класс для хранения соответствия длины текста и размера шрифта.
        /// </summary>
        private class FontOptions
        {        
            /// <summary>
            /// Размер шрифта.
            /// </summary>
            public int Size { get; }

            /// <summary>
            /// Длина входного текста.
            /// </summary>
            public int Length { get; }

            /// <summary>
            /// Перенос текста в ячейке.
            /// </summary>
            public bool Wrap { get; }

            /// <summary>
            /// Создает объект с параметрами шрифта.
            /// </summary>
            /// <param name="size"> Размер шрифта. </param>
            /// <param name="length"> Длина входного текста. </param>
            /// <param name="wrap"> Перенос текста в ячейке. </param>
            public FontOptions(int size, int length, bool wrap)
            {
                Size = size;
                Length = length;
                Wrap = wrap;
            }   
        }
    }
}
