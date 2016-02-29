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

namespace Docx_Various
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // hyperlink
            using (DocX document = DocX.Create(@"Test.docx"))
            {
                // Add a hyperlink to this document.
                Hyperlink h = document.AddHyperlink
                ("Google", new Uri("http://www.google.com"));

                // Add a new Paragraph to this document.
                Paragraph p = document.InsertParagraph();
                p.Append("My favorite search engine is ");
                p.AppendHyperlink(h);
                p.Append(", I think it's great.");

                // Save all changes made to this document.
                document.Save();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {


            // Create a new document.
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Create a new Paragraph with the text "Hello World".
                Paragraph p = document.InsertParagraph("Hello World.");

                // Make this Paragraph flow right to left. Default is left to right.
                //p.Direction = Direction.RightToLeft;
                p.Alignment = Alignment.center;

                // Save all changes made to this document.
                document.Save();
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {


            // Create a new document.
            using (DocX document = DocX.Create("Test.docx"))
            {
                // Create a new Paragraph.
                Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");

                // Indent only the first line of the Paragraph.
                p.IndentationFirstLine = 1.0f;

                // Save all changes made to this document.
                document.Save();
            }


        }

        private void button4_Click(object sender, EventArgs e)
        {


            using (DocX document = DocX.Create("Test.docx"))
            {
                // Add an image to the document.
                Novacode.Image i = document.AddImage(@"Image.jpg");

                // Create a picture i.e. (A custom view of an image)
                Picture p = i.CreatePicture();
                p.FlipHorizontal = true;
                p.Rotation = 10;

                // Create a new Paragraph.
                Paragraph par = document.InsertParagraph();

                // Append content to the Paragraph.
                par.Append("Here is a cool picture")
                   .AppendPicture(p)
                   .Append(" don't you think so?");

                // Save all changes made to this document.
                document.Save();
            }
        }
    }
}
