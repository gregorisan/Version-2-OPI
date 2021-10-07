using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp7
{
    public partial class Form3 : Form
    {
        Word.Application app;
        Word.Document doc;
        int i = 1;
        public Form3()
        {
            InitializeComponent();
        }
        

        private void Form3_Load(object sender, EventArgs e)
        {
            app = new Word.Application();
            doc = app.Documents.Add();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            doc.Paragraphs[1].Range.InsertAfter(textBox1.Text);
            doc.Paragraphs[1].Range.InsertParagraphAfter();
            app.Visible = true;
            i++;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 newForm = new Form1();
            newForm.Show();
            this.Hide();
        }
    }
}
