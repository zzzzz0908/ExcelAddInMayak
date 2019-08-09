using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
//using Office = Microsoft.Office.Core;
//using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        // Some vars
        private const int firstLine = 9;
        //private const string priceList = "Прайс";
        

        #region Накладная

        /// <summary>
        /// Adds line using barcode
        /// </summary>
        /// <param name="text"></param>
        public void AddBarCodeLine(string text, int priceType = 1)
        {            
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string range = $"E{firstLine}:E{firstLine + 500}"; // столбик штрихкодов
            var cell = activeSheet.Range[range];

            foreach (Excel.Range c in cell.Cells)
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
                    c.Offset[0, -4].Value2 = row - firstLine + 1;                    

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
                    c.Offset[1, 4].FormulaLocal = $"=СУММ(I{firstLine}:I{row})";

                    var bigSum = activeSheet.Range["O3"];
                    bigSum.FormulaLocal = $"=I{row + 1}";

                    break;
                }
            }            
        }
        

        /// <summary>
        /// Adds line using article
        /// </summary>
        /// <param name="text"></param>
        /// <param name="priceType"></param>
        public void AddArticleLine(string text, int priceType = 1)
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string range = $"B{firstLine}:B{firstLine + 500}"; // столбик артикулов
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
                    c.Offset[0, -1].Value2 = row - firstLine + 1;

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
                    c.Offset[1, 7].FormulaLocal = $"=СУММ(I{firstLine}:I{row})";

                    var bigSum = activeSheet.Range["O3"];
                    bigSum.FormulaLocal = $"=I{row + 1}";

                    break;
                }
            }
        }
        
        /// <summary>
        /// Deletes last added line
        /// </summary>
        public void DeleteLastLine()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string range = $"E{firstLine}:E{firstLine + 500}";
            var cell = activeSheet.Range[range];

            foreach (Excel.Range c in cell.Cells)
            {
                string value = Convert.ToString(c.Value2);

                if (value == null)
                {
                    int deleteRow = c.Row - 1;
                    if (deleteRow >= firstLine)
                    {
                        range = $"A{c.Row - 1}:L{c.Row - 1}";
                        activeSheet.Range[range].Delete();  // последние строки сами сдвигаются вверх, надо править формулу в сумме?
                    }
                    break;
                }
            }
        }
        
        /// <summary>
        /// Adds bottom of the table
        /// </summary>
        public void AddSummary()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            string range = $"A{firstLine}:A{firstLine + 500}"; // первый столбик
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
                    cells = activeSheet.Range[$"A{firstLine}:I{row - 1}"];
                    cells.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    break;
                }
            }
        }
                
        /// <summary>
        /// Shows barcode column
        /// </summary>
        public void ShowBarCodeColumn()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            var column = activeSheet.Range["E1"];
            column.ColumnWidth = 14;
        }

        /// <summary>
        /// Hides barcode column
        /// </summary>
        public void HideBarCodeColumn()
        {
            var activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            var column = activeSheet.Range["E1"];
            column.ColumnWidth = 0;
        }

        #endregion

        #region Ценники

        /// <summary>
        /// Creates new sheet with price  taags
        /// </summary>
        public void CreatePriceTagSheet()
        {
            this.Application.ScreenUpdating = false;
            var firstSheet = this.Application.ActiveSheet as Excel.Worksheet;            

            Excel.Worksheet newWorksheet;
            newWorksheet = (Excel.Worksheet)this.Application.Worksheets.Add();
            //newWorksheet = CreateWorksheet((Excel.Worksheets)this.Application.Worksheets);
            //Globals.ThisAddIn.Application.ActiveWindow.View = Excel.XlWindowView.xlPageLayoutView;            
            //Thread.Sleep(300);

            // установить ширину столбиков на новом листе
            char[] cols = { 'A', 'E', 'I' };

            foreach (char col in cols)
            {
                var cell = newWorksheet.Range[$"{col}1"];                
                //14.29 2.86 10.14 2.57 ширина
                cell.ColumnWidth = 14.29;
                cell.Offset[0, 1].ColumnWidth = 2.86;
                cell.Offset[0, 2].ColumnWidth = 10.14;
                cell.Offset[0, 3].ColumnWidth = 2.57;
            }                     
             
            var range = firstSheet.Range["A1:A100"];

            // цикл по первому листу
            // for ()
            // перевести число в индекс ячейки
            // получить ячейку с листа
            // создать ценник

            foreach (Excel.Range cell in range.Cells)
            {
                string value = Convert.ToString(cell.Value2);

                if (value != null)
                {
                    int idx = cell.Row;
                    var targetCell = newWorksheet.Range[GetStartCell(idx)];

                    if (idx % 3 == 1)
                    {
                        // 43.50 18.00 28.50 24.00 высота
                        targetCell.RowHeight = 43.50;
                        targetCell.Offset[1, 0].RowHeight = 18.00;
                        targetCell.Offset[2, 0].RowHeight = 28.50;
                        targetCell.Offset[3, 0].RowHeight = 24.00;
                    }
                    // массив данных для ценника: 0 - название, 1 - артикул, 2 - цена опт, 3 - цена розница
                    string[] info = new string[4];
                    info[0] = Convert.ToString(cell.Offset[0, 3].Value2);
                    info[1] = Convert.ToString(cell.Offset[0, 2].Value2);
                    info[2] = Convert.ToString(cell.Offset[0, 5].Value2);
                    info[3] = Convert.ToString(cell.Offset[0, 6].Value2); 

                    CreatePriceTag(targetCell, info);
                }
                else
                {
                    break;
                }
            }

            // настроить поля
            newWorksheet.PageSetup.HeaderMargin = 0;
            newWorksheet.PageSetup.FooterMargin = 0;
            double margin = newWorksheet.Application.CentimetersToPoints(0.5);
            newWorksheet.PageSetup.TopMargin = margin;
            newWorksheet.PageSetup.BottomMargin = margin;
            newWorksheet.PageSetup.LeftMargin = margin;
            newWorksheet.PageSetup.RightMargin = margin;

            this.Application.ScreenUpdating = true;
        }

        /// <summary>
        /// Creates price tag
        /// </summary>
        /// <param name="firstCell"> Upper left cell of price tag</param>
        /// <param name="info"></param>
        private void CreatePriceTag(Excel.Range firstCell, string[] info)
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

        private string GetStartCell(int idx)
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

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
