using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn1
{
    public static class PriceTag
    {
        #region Ценники

        public enum PriceTagSize
        {
            Small = 0,
            Big = 1
        }


        /// <summary>
        /// Creates new sheet with price  tags
        /// </summary>
        /// <param name="tagSize"></param>
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

            switch (tagSize)
            {
                case PriceTagSize.Small:


                    colCount = 3;
                    colsWidth = new double[] { 14.29, 2.86, 10.14, 2.57, 14.29, 2.86, 10.14, 2.57, 14.29, 2.86, 10.14, 2.57 };
                    rowsHeight = new double[] { 43.50, 18.00, 28.50, 24.00 };
                    fontSize = new int[] { 12, 18, 14, 24, 20, 18, 20, 16, 8 };
                    fontOptions = new FontOptions[] {
                        new FontOptions(22, 0, false),
                        new FontOptions(20, 16, false),
                        new FontOptions(18, 19, true),
                        new FontOptions(16, 34, true),
                        new FontOptions(14, 44, true),
                        new FontOptions(12, 56, true) };

                    break;

                case PriceTagSize.Big:
                    colCount = 2;
                    colsWidth = new double[] { 23.14, 6.86, 19.43, 6.86, 23.14, 6.86, 19.43, 6.86 };
                    rowsHeight = new double[] { 57.75, 29.25, 42.75, 28.50 };
                    fontSize = new int[] { 14, 24, 22, 36, 28, 28, 28, 26, 11 };
                    fontOptions = new FontOptions[] {
                        new FontOptions(36, 0, false),
                        new FontOptions(28, 16, false),
                        new FontOptions(24, 22, true),
                        new FontOptions(22, 44, true),
                        new FontOptions(20, 52, true),
                        new FontOptions(18, 60, true),
                        new FontOptions(16, 70, true),
                        new FontOptions(14, 80, true) };
                    break;
            }

            var app = Globals.ThisAddIn.Application;

            app.ScreenUpdating = false;

            var firstSheet = app.ActiveSheet as Excel.Worksheet;
            object[,] values = firstSheet.Range["A1:G100"].Value2;

            Excel.Worksheet newWorksheet = (Excel.Worksheet)app.Worksheets.Add();
            // установить ширину столбиков на новом листе 
            newWorksheet.Range[$"A1:{(char)(colCount * 4 + 64)}1"].ColumnWidth = colsWidth;            


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

#if DEBUG
            stopwatch.Stop();
            newWorksheet.Range["M1"].Value2 = stopwatch.ElapsedMilliseconds;
#endif
        }

        /// <summary>
        /// Creates price tag
        /// </summary>
        /// <param name="firstCell"> Upper left cell of price tag</param>
        /// <param name="info"></param>
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
                       
        #endregion

        private struct FontOptions
        {
            
            public int Size { get; }
            public int Length { get; }
            public bool Wrap { get; }

            public FontOptions(int size, int length, bool wrap)
            {
                Size = size;
                Length = length;
                Wrap = wrap;
            }                        

        }
    }

    
    




}
