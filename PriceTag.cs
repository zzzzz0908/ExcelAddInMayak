using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            int colCount;
            double[] colsWidth = new double[1];

            switch (tagSize)
            {
                case PriceTagSize.Small:
                    colCount = 3;
                    colsWidth = new double[] { 14.29, 2.86, 10.14, 2.57, 14.29, 2.86, 10.14, 2.57, 14.29, 2.86, 10.14, 2.57 };
                    break;
                case PriceTagSize.Big:
                    colCount = 2;
                    colsWidth = new double[] { 14.29, 2.86, 10.14, 2.57, 14.29, 2.86, 10.14, 2.57 };
                    break;
            }

            var app = Globals.ThisAddIn.Application;

            app.ScreenUpdating = false;
            var firstSheet = app.ActiveSheet as Excel.Worksheet;

            Excel.Worksheet newWorksheet = (Excel.Worksheet)app.Worksheets.Add();

            // установить ширину столбиков на новом листе 
            newWorksheet.Range["A1:L1"].ColumnWidth = colsWidth;

            //var range = firstSheet.Range["A1:A100"];

            object[,] values = firstSheet.Range["A1:G100"].Value2;


            for (int i = 1; i <= values.GetLength(0); i++)
            {
                if (values[i, 1] != null)
                {
                    //int idx = cell.Row;
                    var targetCell = newWorksheet.Range[GetStartCell(i)];

                    if (i % 3 == 1)
                    {
                        // 43.50 18.00 28.50 24.00 высота
                        targetCell.RowHeight = 43.50;
                        targetCell.Offset[1, 0].RowHeight = 18.00;
                        targetCell.Offset[2, 0].RowHeight = 28.50;
                        targetCell.Offset[3, 0].RowHeight = 24.00;
                    }
                    // массив данных для ценника: 0 - название, 1 - артикул, 2 - цена опт, 3 - цена розница
                    string[] info = new string[4];
                    info[0] = Convert.ToString(values[i, 4]);
                    info[1] = Convert.ToString(values[i, 3]);
                    info[2] = Convert.ToString(values[i, 6]);
                    info[3] = Convert.ToString(values[i, 7]);

                    CreatePriceTag(targetCell, info);
                }
                else
                {
                    break;
                }
            }

            //foreach (Excel.Range cell in range.Cells)
            //{
            //    string value = Convert.ToString(cell.Value2);

            //    if (value != null)
            //    {
            //        int idx = cell.Row;
            //        var targetCell = newWorksheet.Range[GetStartCell(idx)];

            //        if (idx % 3 == 1)
            //        {
            //            // 43.50 18.00 28.50 24.00 высота
            //            targetCell.RowHeight = 43.50;
            //            targetCell.Offset[1, 0].RowHeight = 18.00;
            //            targetCell.Offset[2, 0].RowHeight = 28.50;
            //            targetCell.Offset[3, 0].RowHeight = 24.00;
            //        }
            //        // массив данных для ценника: 0 - название, 1 - артикул, 2 - цена опт, 3 - цена розница
            //        string[] info = new string[4];
            //        info[0] = Convert.ToString(cell.Offset[0, 3].Value2);
            //        info[1] = Convert.ToString(cell.Offset[0, 2].Value2);
            //        info[2] = Convert.ToString(cell.Offset[0, 5].Value2);
            //        info[3] = Convert.ToString(cell.Offset[0, 6].Value2); 

            //        CreatePriceTag(targetCell, info);
            //    }
            //    else
            //    {
            //        break;
            //    }
            //}

            // настроить поля
            newWorksheet.PageSetup.HeaderMargin = 0;
            newWorksheet.PageSetup.FooterMargin = 0;
            double margin = newWorksheet.Application.CentimetersToPoints(0.5);
            newWorksheet.PageSetup.TopMargin = margin;
            newWorksheet.PageSetup.BottomMargin = margin;
            newWorksheet.PageSetup.LeftMargin = margin;
            newWorksheet.PageSetup.RightMargin = margin;

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
        private static void CreatePriceTag(Excel.Range firstCell, string[] info)
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
            curCell.Font.Size = 18;
            curCell.Font.Bold = true;

            curCell = firstCell.Offset[1, 2];
            curCell.Value2 = "Цена розн.";
            curCell.Font.Size = 14;

            // значок рубля
            firstCell.Offset[2, 1].Value2 = "₽";
            firstCell.Offset[2, 1].Font.Size = 20;

            firstCell.Offset[2, 3].Value2 = "₽";
            firstCell.Offset[2, 3].Font.Size = 20;

            // артикул
            curCell = firstCell.Offset[3, 0];
            curCell.Value2 = info[1];
            curCell.Font.Size = 16;
            curCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            // флп
            firstCell.Offset[3, 2].Value2 = "ФЛП Никулин Ю.В.";
            firstCell.Offset[3, 2].Font.Size = 8;

            // цена опт
            curCell = firstCell.Offset[2, 0];
            curCell.Value2 = double.Parse(info[2]);
            curCell.NumberFormat = "0.00";
            curCell.Font.Size = 24;
            curCell.Font.Bold = true;

            // цена розница
            curCell = firstCell.Offset[2, 2];
            curCell.Value2 = double.Parse(info[3]);
            curCell.NumberFormat = "0.00";
            curCell.Font.Size = 18;

            // объединять первую строчку в конце
            workSheet.Range[firstCell, firstCell.Offset[0, 3]].Merge();

            // название товара
            int nameLength = info[0].Length;
            firstCell.Value2 = info[0];
            firstCell.Font.Size = 12;
            firstCell.WrapText = true;
            firstCell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            if (nameLength <= 56) firstCell.Font.Size = 14;
            if (nameLength <= 44) firstCell.Font.Size = 16;
            if (nameLength <= 34) firstCell.Font.Size = 18;

            if (nameLength <= 19)
            {
                firstCell.Font.Size = 20;
                firstCell.WrapText = false;
            }

            if (nameLength <= 16)
            {
                firstCell.Font.Size = 22;
                firstCell.WrapText = false;
            }

        }

        private static string GetStartCell(int idx)
        {
            // A = 65
            const int width = 4;
            const int height = 4;

            int row = (idx - 1) / 3;
            int col = (idx - 1) % 3;

            char excelCol = (char)(col * width + 65);

            return $"{excelCol}{row * height + 1}";
        }




        #endregion


    }
}
