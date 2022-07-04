using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using ClosedXML.Excel;


namespace ConsoleApp2
{

    interface IFile
    {
        void Write(string Path, string Content);
        string Read(string Path);
    }

    class PDFFile : IFile
    {
        public string Read(string FileName)
        {
            string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + FileName + ".pdf";
            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader(Path))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                }
            }
            return text.ToString();
        }

        public void Write(string FileName, string Content)
        {
            iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.LETTER, 10, 10, 42, 35);
            string PDFpath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + FileName + ".pdf";
            PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(PDFpath, FileMode.Create));
            doc.Open();
            Paragraph paraghraph = new Paragraph();
            paraghraph.Add(Content);
            doc.Add(paraghraph);
            doc.Close();
        }
    }

    class TXTFile : IFile
    {
        public string Read(string FileName)
        {
            string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + FileName + ".txt";
            string Text = File.ReadAllText(Path);
            return Text;
        }

        public void Write(string FileName, string Content)
        {
            string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + FileName + ".txt";
            File.WriteAllText(Path, Content);
        }
    }
    class ExcelFile : IFile
    {
        public string Read(string FileName)
        {
            string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + FileName + ".xlsx";
            var wbook = new XLWorkbook(Path);
            var ws1 = wbook.Worksheet(1);
            var data = ws1.Cell("A1").GetValue<string>();
            return data;
        }

        public void Write(string FileName, string Content)
        {
            string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + FileName + ".xlsx";
            var wbook = new XLWorkbook();

            var ws = wbook.Worksheets.Add("Sheet1");
            ws.Cell("A1").Value = Content;

            wbook.SaveAs(Path);
        }

        //public void Write(string FileName, string Content)
        //{
        //    string Path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "/" + FileName + ".xls";
        //    string connectionString = $"Provider=Microsoft.Jet.OleDb.4.0; Data Source={Path}; Extended Properties=Excel 8.0;";

        //    using (OleDbConnection Connection = new OleDbConnection(connectionString))
        //    {
        //        Connection.Open();
        //        using (OleDbCommand command = new OleDbCommand())
        //        {
        //            command.Connection = Connection;
        //            command.CommandText = $"CREATE TABLE [EmpTable]({Content} Char({Content.Length}))";
        //            command.ExecuteNonQuery();
        //        }
        //    }
        //}


        class DISC
        {
            public List<IFile> Files;
        }

        class Controller
        {
            DISC disc = new DISC();
            public enum FileTypes { PDF = 1, TXT = 2, EXCEL = 3 };
            public void Start()
            {

                while (true)
                {
                    Console.Clear();
                    Console.WriteLine("Read   [1] ");
                    Console.WriteLine("Write  [2] ");
                    int select = int.Parse(Console.ReadLine());
                    if (select == 1)
                    {

                        Console.WriteLine("PDF    [1]");
                        Console.WriteLine("TXT    [2]");
                        Console.WriteLine("EXCEL  [3]");
                        select = int.Parse(Console.ReadLine());
                        if (select == (int)FileTypes.PDF)
                        {
                            Console.WriteLine("Enter File name : ");
                            string filename = Console.ReadLine();
                            PDFFile pdffile = new PDFFile();
                            string content = pdffile.Read(filename);
                            Console.WriteLine($"content : {content}");
                        }
                        else if (select == (int)FileTypes.TXT)
                        {
                            Console.WriteLine("Enter File name : ");
                            string filename = Console.ReadLine();
                            TXTFile txtfile = new TXTFile();
                            string content = txtfile.Read(filename);
                            Console.WriteLine($"content : {content}");
                        }
                        else if (select == (int)FileTypes.EXCEL)
                        {
                            Console.WriteLine("Enter File name : ");
                            string filename = Console.ReadLine();
                            ExcelFile xlsfile = new ExcelFile();
                            string content = xlsfile.Read(filename);
                            Console.WriteLine($"content : {content}");
                        }
                    }
                    //Write Section
                    else if (select == 2)
                    {
                        Console.WriteLine("PDF    [1]");
                        Console.WriteLine("TXT    [2]");
                        Console.WriteLine("EXCEL  [3]");
                        select = int.Parse(Console.ReadLine());

                        if (select == (int)FileTypes.PDF)
                        {
                            Console.WriteLine("Enter File name : ");
                            string filename = Console.ReadLine();
                            Console.WriteLine("Enter Content : ");
                            string content = Console.ReadLine();
                            PDFFile pdffile = new PDFFile();
                            pdffile.Write(filename, content);
                            Console.WriteLine("Writed Succesfully");
                            disc.Files.Add(pdffile);
                        }
                        else if (select == (int)FileTypes.TXT)
                        {
                            Console.WriteLine("Enter File name : ");
                            string filename = Console.ReadLine();
                            Console.WriteLine("Enter Content : ");
                            string content = Console.ReadLine();
                            TXTFile txtfile = new TXTFile();
                            txtfile.Write(filename, content);
                            Console.WriteLine("Writed Succesfully");
                            disc.Files.Add(txtfile);

                        }
                        else if (select == (int)FileTypes.EXCEL)
                        {
                            Console.WriteLine("Enter File name : ");
                            string filename = Console.ReadLine();
                            Console.WriteLine("Enter Content : ");
                            string content = Console.ReadLine();
                            ExcelFile xlsfile = new ExcelFile();
                            xlsfile.Write(filename, content);
                            Console.WriteLine("Writed Succesfully");
                            disc.Files.Add(xlsfile);
                        }
                    }
                    else
                    {
                        continue;
                    }
                    Console.ReadKey();
                }

            }
        }


        public class Program
        {
            static void Main(string[] args)
            {
                Controller controller = new Controller();
                controller.Start();
            }
        }
    }
}
