using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Novacode;
using System.Diagnostics;
namespace testDocx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreateSimpleDocument();
            // ref: http://johnatten.com/2013/09/28/c-create-and-manipulate-word-documents-programmatically-using-docx/
        }

        private void CreateSimpleDocument()
        {
            string filename = textBox2.Text;
            var doc = DocX.Create(filename);
            doc.InsertParagraph(textBox1.Text);
            doc.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start("WINWORD.EXE", textBox2.Text);
        }
    }
}
