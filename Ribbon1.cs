using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            var inputForm = new InputForm() { inputType = InputForm.InputType.BarCode };
            inputForm.Show();
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            var inputForm = new InputForm() { inputType = InputForm.InputType.Article, Text = "Ввод артикула" };            
            inputForm.Show();            
        }

        private void Button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ShowBarCodeColumn();
        }

        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.HideBarCodeColumn();
        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddSummary();
        }

        private void Button6_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CreatePriceTagSheet(ThisAddIn.PriceTagSize.Small);
        }
    }
}
