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


namespace testDocx_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = textBox1.Text;
            string headlineText = "Constitution of the United States";
            string paraOne = ""
                + "We the People of the United States, in Order to form a more perfect Union, "
                + "establish Justice, insure domestic Tranquility, provide for the common defence, "
                + "promote the general Welfare, and secure the Blessings of Liberty to ourselves "
                + "and our Posterity, do ordain and establish this Constitution for the United States of America.";

            // A formatting object for our headline:
            var headLineFormat = new Formatting();
            headLineFormat.FontFamily = new System.Drawing.FontFamily("Arial Black");
            headLineFormat.Size = 18D;
            headLineFormat.Position = 12;



            // A formatting object for our normal paragraph text:
            var paraFormat = new Formatting();
            paraFormat.FontFamily = new System.Drawing.FontFamily("Calibri");
            paraFormat.Size = 10D;

            // Create the document in memory:
            var doc = DocX.Create(fileName);



            // Insert the Headline and do some formatting:
            Paragraph headline = doc.InsertParagraph(headlineText);
            headline.Color(System.Drawing.Color.Blue);
            headline.Font(new System.Drawing.FontFamily("Comic Sans MS"));
            headline.Bold();
            headline.Position(12D);
            headline.FontSize(18D);


            // Insert the now text obejcts;
            doc.InsertParagraph(headlineText, false, headLineFormat);
            doc.InsertParagraph(paraOne, false, paraFormat);



            // Save to the output directory:
            doc.Save();

            // Open in Word:
            Process.Start("WINWORD.EXE", fileName);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CreateRejectionLetter("%APPLICANT%", "Koson");
        }
        public void CreateRejectionLetter(string applicantField, string applicantName)
        {
            // We will need a file name for our output file (change to suit your machine):
            string fileNameTemplate = @"Rejection-Letter-{0}-{1}.docx";

            // Let's save the file with a meaningful name, including the applicant name and the letter date:
            string outputFileName = string.Format(fileNameTemplate, applicantName, DateTime.Now.ToString("MM-dd-yy"));

            // Grab a reference to our document template:
            DocX letter = this.GetRejectionLetterTemplate();

            // Perform the replace:
            letter.ReplaceText(applicantField, applicantName);

            // Save as New filename:
            letter.SaveAs(outputFileName);

            // Open in word:
            Process.Start("WINWORD.EXE", "\"" + outputFileName + "\"");
        }
        private DocX GetRejectionLetterTemplate()
        {
            // Adjust the path so suit your machine:
            string fileName = textBox1.Text;

            // Set up our paragraph contents:
            string headerText = "Rejection Letter";
            string letterBodyText = DateTime.Now.ToShortDateString();
            string paraTwo = ""
                + "Dear %APPLICANT%" + Environment.NewLine + Environment.NewLine
                + "I am writing to thank you for your resume. Unfortunately, your skills and "
                + "experience do not match our needs at the present time. We will keep your "
                + "resume in our circular file for future reference. Don't call us, we'll call you. "
                + Environment.NewLine + Environment.NewLine
                + "Sincerely, "
                + Environment.NewLine + Environment.NewLine
                + "Jim Smith, Corporate Hiring Manager";

            // Title Formatting:
            var titleFormat = new Formatting();
            titleFormat.FontFamily = new System.Drawing.FontFamily("Arial Black");
            titleFormat.Size = 18D;
            titleFormat.Position = 12;

            // Body Formatting
            var paraFormat = new Formatting();
            paraFormat.FontFamily = new System.Drawing.FontFamily("Calibri");
            paraFormat.Size = 10D;
            titleFormat.Position = 12;

            // Create the document in memory:
            var doc = DocX.Create(fileName);

            // Insert each prargraph, with appropriate spacing and alignment:
            Paragraph title = doc.InsertParagraph(headerText, false, titleFormat);
            title.Alignment = Alignment.center;

            doc.InsertParagraph(Environment.NewLine);
            Paragraph letterBody = doc.InsertParagraph(letterBodyText, false, paraFormat);
            letterBody.Alignment = Alignment.both;

            doc.InsertParagraph(Environment.NewLine);
            doc.InsertParagraph(paraTwo, false, paraFormat);

            return doc;

        }
    }
}
