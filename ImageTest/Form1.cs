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


namespace ImageTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Create a .docx file
            using (DocX document = DocX.Create(@"Example.docx"))
            {
                // Add an Image to the docx file
                Novacode.Image img = document.AddImage(
                @"donkey.jpg");

                // Insert an emptyParagraph into this document.
                Paragraph p = document.InsertParagraph("", false);

                #region pic1
                Picture pic1 = img.CreatePicture();

                // Set the Picture pic1’s shape
                pic1.SetPictureShape(BasicShapes.cube);

                // Rotate the Picture pic1 clockwise by 30 degrees
                pic1.Rotation = 30;
                p.InsertPicture(pic1, 0);

                #endregion

                // Save the docx file
                document.Save();
            }
        }
    }
}
