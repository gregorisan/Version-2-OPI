using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp7
{
    public partial class Form2 : Form
    {

        Excel.Workbook workBook;
        Excel.Worksheet workSheet;
        Excel.Application excelApp;
        int i = 1;

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            excelApp = new Excel.Application();


            workBook = excelApp.Workbooks.Add();
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 newForm = new Form1();
            newForm.Show();
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!excelApp.Visible == true) excelApp.Visible = true;
            if (!excelApp.UserControl == true) excelApp.UserControl = true;

            workSheet.Cells[Convert.ToInt32(textBox3.Text),textBox2.Text] = textBox1.Text;

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = Convert.ToString(workSheet.Cells[Convert.ToInt32(textBox3.Text), textBox2.Text].Text);
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
