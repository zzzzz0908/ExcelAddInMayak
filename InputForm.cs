using System;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn1
{
    

    public partial class InputForm : Form
    {
        public enum InputType
        {
            BarCode = 0,
            Article = 1
        }


        public InputType inputType { get; set; }

        public InputForm()
        {
            InitializeComponent();
        }


        private void TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\u001B')
            {
                Invoice.DeleteLastLine();
                e.Handled = true;
            }
            else if (e.KeyChar == '\r')
            {
                TextBox tb = (TextBox)sender;
                string text = tb.Text;

                if (tb.Text != "")
                {
                    int priceType = 1;

                    foreach (RadioButton rb in groupBox1.Controls)
                    {
                        if (rb.Checked)
                        {
                            priceType = Convert.ToInt32(rb.Tag);
                        }
                    }

                    switch (inputType)
                    {
                        case InputType.BarCode:
                            Invoice.AddBarCodeLine(text, priceType);
                            break;
                        case InputType.Article:
                            Invoice.AddArticleLine(text, priceType);
                            break;
                    }
                }

                //Globals.ThisAddIn.AddLine(text, priceType);

                tb.Clear();
                e.Handled = true;
            }
        }

        private void ActivateTextBox(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            textBox1.Clear();            
        }

        
    }
}
