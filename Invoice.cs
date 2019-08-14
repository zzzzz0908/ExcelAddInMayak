﻿using System;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn1
{
    public static class Invoice
    {      
        // Константы
        private const int FIRST_LINE = 9;

        /// <summary>
        /// Добавляет запись по штрихкоду.
        /// </summary>
        /// <param name="text"> Штрихкод. </param>
        /// <param name="priceType"> Тип цены. </param>
        public static void AddBarCodeLine(string text, int priceType = 1)
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            // столбик штрихкодов
            string rangeName = $"E{FIRST_LINE}:E{FIRST_LINE + 500}";
            
            var range = activeSheet.Range[rangeName];

            foreach (Excel.Range c in range.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (text == value)
                {
                    // добавить проверку, что есть число в количкстве ?
                    c.Offset[0, 1].Value2++;
                    break;
                }
                else if (value == null)
                {
                    c.Value2 = text;
                    c.Offset[0, 1].Value2 = 1;

                    int row = c.Row;
                    c.Offset[0, -4].Value2 = row - FIRST_LINE + 1;

                    c.Offset[0, -3].FormulaLocal = $"=ВПР(E{row};Price!$B$2:$G$7000;2;ЛОЖЬ)";
                    c.Offset[0, -2].FormulaLocal = $"=ВПР(E{row};Price!$B$2:$G$7000;3;ЛОЖЬ)";
                    c.Offset[0, 2].FormulaLocal = $"=ВПР(E{row};Price!$B$2:$G$7000;{4 + priceType};ЛОЖЬ)";
                    c.Offset[0, -1].Value2 = "шт.";

                    if (priceType == 0)
                    {
                        c.Offset[0, 6].Value2 = 0;
                    }
                    else
                    {
                        // поля про скидку
                        c.Offset[0, 7].FormulaLocal = $"=ВПР(E{row};Price!$B$2:$H$7000;7;ЛОЖЬ)";
                        c.Offset[0, 8].FormulaLocal = $"=ВПР(E{row};Price!$B$2:$I$7000;8;ЛОЖЬ)";
                        c.Offset[0, 6].FormulaLocal = $"=ЕСЛИ(L{row}=\"Да\";ЕСЛИ(M{row}<$M$2; M{row};$M$2); 0)";
                    }

                    c.Offset[0, 3].FormulaLocal = $"=G{row}*(1-K{row}/100)";

                    c.Offset[0, 4].FormulaLocal = $"=H{row}*F{row}";

                    c.Offset[1, 3].Value2 = "Итого:";
                    //c.Offset[1, 2].Font.Bold = true;
                    c.Offset[1, 4].FormulaLocal = $"=СУММ(I{FIRST_LINE}:I{row})";

                    var bigSum = activeSheet.Range["O3"];
                    bigSum.FormulaLocal = $"=I{row + 1}";

                    break;
                }
            }
        }


        /// <summary>
        /// Добавляет запись по артикулу.
        /// </summary>
        /// <param name="text"> Артикул. </param>
        /// <param name="priceType"> Тип цены. </param>
        public static void AddArticleLine(string text, int priceType = 1)
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string range = $"B{FIRST_LINE}:B{FIRST_LINE + 500}"; // столбик артикулов
            var cell = activeSheet.Range[range];

            foreach (Excel.Range c in cell.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (text == value)
                {
                    // добавить проверку, что есть число в количкстве ?
                    c.Offset[0, 4].Value2++;
                    break;
                }
                else if (value == null)
                {
                    c.Value2 = text;

                    int row = c.Row;
                    c.Offset[0, -1].Value2 = row - FIRST_LINE + 1;

                    c.Offset[0, 1].FormulaLocal = $"=ВПР(B{row};Price!$C$2:$G$7000;2;ЛОЖЬ)";
                    c.Offset[0, 2].Value2 = "шт.";

                    c.Offset[0, 3].Value2 = -1;
                    c.Offset[0, 4].Value2 = 1;
                    c.Offset[0, 5].FormulaLocal = $"=ВПР(B{row};Price!$C$2:$G$7000;{3 + priceType};ЛОЖЬ)";

                    if (priceType == 0)
                    {
                        c.Offset[0, 9].Value2 = 0;
                    }
                    else
                    {
                        // поля про скидку
                        c.Offset[0, 10].FormulaLocal = $"=ВПР(B{row};Price!$C$2:$H$7000;6;ЛОЖЬ)";
                        c.Offset[0, 11].FormulaLocal = $"=ВПР(B{row};Price!$C$2:$I$7000;7;ЛОЖЬ)";
                        c.Offset[0, 9].FormulaLocal = $"=ЕСЛИ(L{row}=\"Да\";ЕСЛИ(M{row}<$M$2; M{row};$M$2); 0)";
                    }

                    c.Offset[0, 6].FormulaLocal = $"=G{row}*(1-K{row}/100)";

                    c.Offset[0, 7].FormulaLocal = $"=H{row}*F{row}";

                    c.Offset[1, 6].Value2 = "Итого:";
                    //c.Offset[1, 2].Font.Bold = true;
                    c.Offset[1, 7].FormulaLocal = $"=СУММ(I{FIRST_LINE}:I{row})";

                    var bigSum = activeSheet.Range["O3"];
                    bigSum.FormulaLocal = $"=I{row + 1}";

                    break;
                }
            }
        }

        /// <summary>
        /// Удалаяет последнюю щапись.
        /// </summary>
        public static void DeleteLastLine()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string range = $"E{FIRST_LINE}:E{FIRST_LINE + 500}";
            var cell = activeSheet.Range[range];

            foreach (Excel.Range c in cell.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (value == null)
                {
                    int deleteRow = c.Row - 1;
                    if (deleteRow >= FIRST_LINE)
                    {
                        range = $"A{c.Row - 1}:L{c.Row - 1}";
                        activeSheet.Range[range].Delete();  // последние строки сами сдвигаются вверх, надо править формулу в сумме?
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// Добавляет итог накладной.
        /// </summary>
        public static void AddSummary()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string range = $"A{FIRST_LINE}:A{FIRST_LINE + 500}"; // первый столбик
            var cells = activeSheet.Range[range];

            foreach (Excel.Range c in cells.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (value == null)
                {
                    int row = c.Row;

                    c.Offset[2, 0].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    c.Offset[2, 0].WrapText = false;
                    c.Offset[2, 0].Value2 = "Всего на сумму:";

                    c.Offset[3, 0].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    c.Offset[3, 0].WrapText = false;
                    c.Offset[3, 0].FormulaLocal = $"=РосРуб(I{row};I{row})";

                    c.Offset[5, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    c.Offset[5, 1].WrapText = false;
                    c.Offset[5, 1].Value2 = "Отгрузил(а)";

                    c.Offset[5, 3].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    c.Offset[5, 3].WrapText = false;
                    c.Offset[5, 3].Value2 = "Получил(а)";

                    // толстые линии
                    c.Offset[5, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[5, 2].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;

                    c.Offset[5, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[5, 8].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                    c.Offset[5, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[5, 6].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;
                    c.Offset[5, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[5, 7].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlMedium;


                    // граница итого
                    c.Offset[0, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[0, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    //граница товаров
                    cells = activeSheet.Range[$"A{FIRST_LINE}:I{row - 1}"];
                    cells.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    break;
                }
            }
        }

        /// <summary>
        /// Показывает столбец со штрихкодом.
        /// </summary>
        public static void ShowBarCodeColumn()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            var column = activeSheet.Range["E1"];
            column.ColumnWidth = 14;
        }

        /// <summary>
        /// Скрывает столбец со штрихкодом.
        /// </summary>
        public static void HideBarCodeColumn()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            var column = activeSheet.Range["E1"];
            column.ColumnWidth = 0;
        }        
    }
}
