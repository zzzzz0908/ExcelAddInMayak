using System;
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

            int inputCol = 5;
            string valueSheet = "Price";
            
            foreach (Excel.Range c in range.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (text == value)
                {
                    // добавить проверку, что есть число в количкстве ?
                    c.Offset[0, 7 - inputCol].Value2++;
                    break;
                }
                else if (value == null)
                {
                    int row = c.Row;

                    c.Value2 = text;              
                    
                    string inputCellName = GetRangeName(row, inputCol);

                    c.Offset[0, 1 - inputCol].Value2 = row - FIRST_LINE + 1;
                    c.Offset[0, 2 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$B$2:$G$7000;2;ЛОЖЬ)";
                    c.Offset[0, 3 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$B$2:$G$7000;3;ЛОЖЬ)";
                    c.Offset[0, 4 - inputCol].Value2 = "шт.";

                    c.Offset[0, 6 - inputCol].Value2 = 1;
                    c.Offset[0, 7 - inputCol].Value2 = 1;
                    c.Offset[0, 8 - inputCol].FormulaLocal = $"=F{row}*G{row}";
                    c.Offset[0, 9 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$B$2:$G$7000;{4 + priceType};ЛОЖЬ)";                    

                    if (priceType == 0)
                    {
                        c.Offset[0, 13 - inputCol].Value2 = 0;
                    }
                    else
                    {
                        // поля про скидку
                        c.Offset[0, 16 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$B$2:$H$7000;7;ЛОЖЬ)";
                        c.Offset[0, 17 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$B$2:$I$7000;8;ЛОЖЬ)";
                        c.Offset[0, 15 - inputCol].FormulaLocal = $"=ЕСЛИ(P{row}=\"Да\";ЕСЛИ(Q{row}<$N$2; Q{row};$N$2); 0)";
                    }

                    c.Offset[0, 10 - inputCol].FormulaLocal = $"=ОКРУГЛ(I{row}*(1-O{row}/100);2)";
                    c.Offset[0, 11 - inputCol].FormulaLocal = $"=H{row}*I{row}";
                    c.Offset[0, 12 - inputCol].FormulaLocal = $"=K{row}-M{row}";
                    c.Offset[0, 13 - inputCol].FormulaLocal = $"=J{row}*H{row}";


                    c.Offset[1, 10 - inputCol].Value2 = "Итого:";
                    c.Offset[1, 11 - inputCol].FormulaLocal = $"=СУММ(K{FIRST_LINE}:K{row})";
                    c.Offset[1, 12 - inputCol].FormulaLocal = $"=СУММ(L{FIRST_LINE}:L{row})";
                    c.Offset[1, 13 - inputCol].FormulaLocal = $"=СУММ(M{FIRST_LINE}:M{row})";

                    var bigSum = activeSheet.Range["O1"];
                    bigSum.FormulaLocal = "=" + GetRangeName(row + 1, 13);

                    // прокрутка экрана
                    c.Offset[1, 1 - inputCol].Activate();

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

            int inputCol = 2;
            string valueSheet = "Price";

            foreach (Excel.Range c in cell.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (text == value)
                {
                    // добавить проверку, что есть число в количкстве ?
                    c.Offset[0, 7 - inputCol].Value2++;
                    break;
                }
                else if (value == null)
                {
                    int row = c.Row;
                    c.Value2 = text;
                    string inputCellName = GetRangeName(row, inputCol);

                    c.Offset[0, 1 - inputCol].Value2 = row - FIRST_LINE + 1;

                    c.Offset[0, 3 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$C$2:$G$7000;2;ЛОЖЬ)";
                    c.Offset[0, 4 - inputCol].Value2 = "шт.";
                    c.Offset[0, 5 - inputCol].Value2 = -1;
                    c.Offset[0, 6 - inputCol].Value2 = 1;
                    c.Offset[0, 7 - inputCol].Value2 = 1;
                    c.Offset[0, 8 - inputCol].FormulaLocal = $"=F{row}*G{row}";
                    c.Offset[0, 9 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$C$2:$G$7000;{3 + priceType};ЛОЖЬ)";

                    if (priceType == 0)
                    {
                        c.Offset[0, 13 - inputCol].Value2 = 0;
                    }
                    else
                    {
                        // поля про скидку
                        c.Offset[0, 16 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$C$2:$H$7000;6;ЛОЖЬ)";
                        c.Offset[0, 17 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$C$2:$I$7000;7;ЛОЖЬ)";
                        c.Offset[0, 15 - inputCol].FormulaLocal = $"=ЕСЛИ(P{row}=\"Да\";ЕСЛИ(Q{row}<$N$2; Q{row};$N$2); 0)";
                    }

                    c.Offset[0, 10 - inputCol].FormulaLocal = $"=ОКРУГЛ(I{row}*(1-O{row}/100);2)";
                    c.Offset[0, 11 - inputCol].FormulaLocal = $"=H{row}*I{row}";
                    c.Offset[0, 12 - inputCol].FormulaLocal = $"=K{row}-M{row}";
                    c.Offset[0, 13 - inputCol].FormulaLocal = $"=J{row}*H{row}";


                    c.Offset[1, 10 - inputCol].Value2 = "Итого:";
                    c.Offset[1, 11 - inputCol].FormulaLocal = $"=СУММ(K{FIRST_LINE}:K{row})";
                    c.Offset[1, 12 - inputCol].FormulaLocal = $"=СУММ(L{FIRST_LINE}:L{row})";
                    c.Offset[1, 13 - inputCol].FormulaLocal = $"=СУММ(M{FIRST_LINE}:M{row})";

                    var bigSum = activeSheet.Range["O1"];
                    bigSum.FormulaLocal = "=" + GetRangeName(row + 1, 13);

                    // прокрутка экрана
                    c.Offset[1, 1 - inputCol].Activate();

                    break;
                }
            }
        }


        /// <summary>
        /// Добавляет запись по штрихкоду упаковки (ящика).
        /// </summary>
        /// <param name="text"> Штрихкод. </param>
        /// <param name="priceType"> Тип цены. </param>
        public static void AddBoxLine(string text, int priceType = 1)
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            // столбик штрихкодов
            string rangeName = $"E{FIRST_LINE}:E{FIRST_LINE + 500}";

            var range = activeSheet.Range[rangeName];

            int inputCol = 5;
            string valueSheet = "BoxPrice";

            foreach (Excel.Range c in range.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (text == value)
                {
                    // добавить проверку, что есть число в количкстве ?
                    c.Offset[0, 7 - inputCol].Value2++;
                    break;
                }
                else if (value == null)
                {
                    int row = c.Row;

                    c.Value2 = text;

                    string inputCellName = GetRangeName(row, inputCol);

                    c.Offset[0, 1 - inputCol].Value2 = row - FIRST_LINE + 1;
                    c.Offset[0, 2 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$A$2:$K$7000;5;ЛОЖЬ)";
                    c.Offset[0, 3 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$A$2:$K$7000;6;ЛОЖЬ)";
                    c.Offset[0, 4 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$A$2:$K$7000;3;ЛОЖЬ)";

                    c.Offset[0, 6 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$A$2:$K$7000;2;ЛОЖЬ)";
                    c.Offset[0, 7 - inputCol].Value2 = 1;
                    c.Offset[0, 8 - inputCol].FormulaLocal = $"=F{row}*G{row}";
                    c.Offset[0, 9 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$A$2:$K$7000;{7 + priceType};ЛОЖЬ)";

                    if (priceType == 0)
                    {
                        c.Offset[0, 13 - inputCol].Value2 = 0;
                    }
                    else
                    {
                        // поля про скидку
                        c.Offset[0, 16 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$A$2:$K$7000;10;ЛОЖЬ)";
                        c.Offset[0, 17 - inputCol].FormulaLocal = $"=ВПР({inputCellName};{valueSheet}!$A$2:$K$7000;11;ЛОЖЬ)";
                        c.Offset[0, 15 - inputCol].FormulaLocal = $"=ЕСЛИ(P{row}=\"Да\";ЕСЛИ(Q{row}<$N$2; Q{row};$N$2); 0)";
                    }

                    c.Offset[0, 10 - inputCol].FormulaLocal = $"=ОКРУГЛ(I{row}*(1-O{row}/100);2)";
                    c.Offset[0, 11 - inputCol].FormulaLocal = $"=H{row}*I{row}";
                    c.Offset[0, 12 - inputCol].FormulaLocal = $"=K{row}-M{row}";
                    c.Offset[0, 13 - inputCol].FormulaLocal = $"=J{row}*H{row}";


                    c.Offset[1, 10 - inputCol].Value2 = "Итого:";
                    c.Offset[1, 11 - inputCol].FormulaLocal = $"=СУММ(K{FIRST_LINE}:K{row})";
                    c.Offset[1, 12 - inputCol].FormulaLocal = $"=СУММ(L{FIRST_LINE}:L{row})";
                    c.Offset[1, 13 - inputCol].FormulaLocal = $"=СУММ(M{FIRST_LINE}:M{row})";

                    var bigSum = activeSheet.Range["O1"];
                    bigSum.FormulaLocal = "=" + GetRangeName(row + 1, 13);

                    // прокрутка экрана
                    c.Offset[1, 1 - inputCol].Activate();

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
                    c.Offset[3, 0].FormulaLocal = $"=РосРуб(K{row};K{row})";

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
                    c.Offset[0, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[0, 10].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[0, 11].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    c.Offset[0, 12].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


                    //граница товаров
                    cells = activeSheet.Range[$"A{FIRST_LINE}:{GetRangeName(row - 1, 13)}"];
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


        public static void ChangePriceType(int priceType)
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string rangeName = $"A{FIRST_LINE}:A{FIRST_LINE + 500}";

            var range = activeSheet.Range[rangeName];

            //int inputCol = 5;
            //string valueSheet = "BoxPrice";

            foreach (Excel.Range c in range.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (value == null)
                {
                    break;
                }
                else
                {
                    //int row = c.Row;
                    string formula = c.Offset[0, 8].FormulaLocal;
                    string[] formulaParams = formula.Split(';');
                    string priceRangeName = formulaParams[1];
                    char startColumnLetter = priceRangeName.Substring(priceRangeName.IndexOf("!$") + 2, 1).ToCharArray()[0];

                    
                    // ввод по штрихкоду
                    if (startColumnLetter == 'B')
                    {
                        formulaParams[2] = (4 + priceType).ToString(); // 4 - magic number (сдвиг в ВПР)
                        string newFormula = formulaParams[0] + ";" + formulaParams[1] + ";" + formulaParams[2] + ";" + formulaParams[3];
                        c.Offset[0, 8].FormulaLocal = newFormula;
                        continue;
                    }

                    // ввод по артикулу
                    if (startColumnLetter == 'C')
                    {
                        formulaParams[2] = (3 + priceType).ToString(); // 3 - magic number
                        string newFormula = formulaParams[0] + ";" + formulaParams[1] + ";" + formulaParams[2] + ";" + formulaParams[3];
                        c.Offset[0, 8].FormulaLocal = newFormula;
                        continue;
                    }

                    // ввод по штрихкоду упаковки
                    if (startColumnLetter == 'A')
                    {
                        formulaParams[2] = (7 + priceType).ToString(); // 7 - magic number
                        string newFormula = formulaParams[0] + ";" + formulaParams[1] + ";" + formulaParams[2] + ";" + formulaParams[3];
                        c.Offset[0, 8].FormulaLocal = newFormula;
                        continue;
                    }
                }
            }
        }


                
        private static string GetRangeName(int row, int col)
        {
            var colChar = (char)(col + 64);
            return colChar.ToString() + row.ToString();
        }


    }
}
