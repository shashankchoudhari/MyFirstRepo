using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace OpenXML_Manipulation
{
    class Program
    {
        public void HelloWorld(string docName)
        {
            try
            {
                //This code is written by ayknijA
                //hello Shashank : Trying another way of using git
                // Create a Wordprocessing document. 

                //new comment added

                //Being updated in timely manner.!!
                using (WordprocessingDocument package = WordprocessingDocument.Create(docName, WordprocessingDocumentType.Document))
                {
                    // Add a new main document part. 
                    package.AddMainDocumentPart();

                    // Create the Document DOM. 
                    package.MainDocumentPart.Document =
                      new Document(
                        new Body(
                          new Paragraph(
                            new Run(
                              new Text("Hello World! I am Visual Studio writing in word document!!"))),
                          new Paragraph(
                              new Run(
                                  new Text("Second Line"))),
                          new Paragraph(
                              new Run(
                                  new Text("Third Line")))));



                    // Save changes to the main document part. 
                    package.MainDocumentPart.Document.Save();
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public void ReadWorld(string docName)
        {
            try
            { 
                using (WordprocessingDocument package = WordprocessingDocument.Open(docName, true))
                {
                    List<Paragraph> paragraphs = package.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList();
                    foreach (var paragraph in paragraphs)
                    {
                        Console.WriteLine(paragraph.InnerText);
                    }
                }
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void Main(string[] args)
        {
            //Program pgm = new Program();
            //pgm.HelloWorld(@"D:\OpenXml Docs\SampleDoc.docx");
            //pgm.ReadWorld(@"D:\OpenXml Docs\SampleDoc.docx");

            StopwatchProcessing.Estimate();
        }
    }
}
