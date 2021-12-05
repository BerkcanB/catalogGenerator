/* This program created by Berkcan Bilçer. See in github "BerkcanB"*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDataReader;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;



namespace eCatalog
{
    public partial class MainScreen : Form
    {
        public MainScreen()
        {
            InitializeComponent();
        }

        private void MainScreen_Load(object sender, EventArgs e)
        {
            CreatePDF(OpenExcel());
        }

        DataTable OpenExcel()
        {
            var stream = File.Open("Excel_Document.xlsx", FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();
            DataTable dTable = new DataTable();
            foreach (DataTable table in tables)
            {
                dTable = table;
            }//1 0 isim 1 1 aciklama 1 2 fiyat 1 3 resim adi
            stream.Close();
            return dTable;
        }

        void CreatePDF(DataTable table)
        {
            int pageNumber = 1;
            PdfDocument coverDoc = PdfReader.Open("docs//Cover-Example.pdf", PdfDocumentOpenMode.Modify);//Cover of Pdf document
            while (table.Rows.Count>pageNumber)//It does for all rows
            {
                PdfDocument document = PdfReader.Open("docs//Page-Example.pdf", PdfDocumentOpenMode.Modify);
                document.Info.Title = "Berkcan PDF Printer";
                PdfPage page = document.Pages[0];
                XGraphics gfx = XGraphics.FromPdfPage(page);
                XFont font = new XFont("Times New Roman", 30, XFontStyle.Regular);
                gfx.DrawString(table.Rows[pageNumber][0].ToString(), //Name
                                          font, XBrushes.Black, //font and brush style
                                          new XRect(0, 70, page.Width, page.Height), //location
                                          XStringFormats.Center); //location of location's head

                gfx.DrawString(table.Rows[pageNumber][1].ToString(),//Explanation
                                          font, XBrushes.Black,
                                          new XRect(0, 200, page.Width, page.Height),
                                          XStringFormats.Center);
                gfx.DrawString(table.Rows[pageNumber][2].ToString()+"$",//Price
                                          font, XBrushes.Black,
                                          new XRect(200, 0, page.Width, page.Height),
                                          XStringFormats.Center);

                gfx.DrawString(pageNumber.ToString(),//Page number
                               font, XBrushes.White,
                               new XRect(-10, -10, page.Width, page.Height),
                               XStringFormats.BottomRight);

                XPoint imagePoint = new XPoint();//image's location
                imagePoint.X = 225; //x coordinate
                imagePoint.Y = 100; //y coordinate
                string imgPath = "img//" + table.Rows[pageNumber][3].ToString();//image path where it is
                XImage img = XImage.FromFile(imgPath);
                double wRatio = (double)img.PixelHeight / (double)img.PixelWidth; //image width and height ratio
                gfx.DrawImage(img, new XRect(new XPoint(150, 100), new XSize(300, wRatio * 300)));
                const string filename = "Page.pdf";//Paper
                document.Save(filename);
                document.Close();
                PdfDocument add = PdfReader.Open("Page.pdf", PdfDocumentOpenMode.Import);
                coverDoc.AddPage(add.Pages[0]);
                File.Delete("Page.pdf");
                add.Close();
                coverDoc.Save("New.pdf");
                pageNumber++;
            }
            coverDoc.Close();

            Process.Start("New.pdf");//Starts the document
            this.Close();
        }
    }
}
