using System;
using System.Xml;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelHelper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filePath;
        string URL = "https://geocode-maps.yandex.ru/1.x/?apikey=4f98a5df-1522-46aa-8279-2a8691188448&geocode=";
        private void btn_OpenFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;
                    tb_Path.Text = filePath;
                }
            }

        }

        private void btn_Calc_Click(object sender, EventArgs e)
        {
            if (filePath != null) 
            {
                Excel.Application ObjExcel = new Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Workbook ObjWorkBook = ObjExcel.Workbooks.Open(filePath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Worksheet ObjWorkSheet;
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

                ObjExcel.Visible = true;

                Range last = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
                Range range = ObjWorkSheet.get_Range("A2", last);

                ObjWorkSheet.Cells[1, 18].Value2 = "Долгота";
                ObjWorkSheet.Cells[1, 19].Value2 = "Широта";

                //Iterate the rows in the used range
                foreach (Excel.Range row in range.Rows)
                { 
                    string result = GetCord(URL + row.Columns[2].Value2 + ", " + row.Columns[5].Value2);
                    string[] values = result.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    row.Columns[18].Value2 = values[0];
                    row.Columns[19].Value2 = values[1];

                }
            }
        }
        private string GetCord(string Http)
        {
            string l = "<Point xmlns=\"http://www.opengis.net/gml\"><pos>";
            WebRequest req = WebRequest.Create(@"" + Http);
            WebResponse resp = req.GetResponse();
            Stream stream = resp.GetResponseStream();
            StreamReader reader = new StreamReader(stream);
            string text = reader.ReadToEnd();
            text = text.Substring(text.IndexOf(l) + l.Length);
            text = text.Remove(text.IndexOf("</pos>"));
            return text;
        }

        private string GetCord2(string Http)
        {
            WebRequest req = WebRequest.Create(@"" + Http);
            WebResponse resp = req.GetResponse();
            Stream stream = resp.GetResponseStream();
            StreamReader reader = new StreamReader(stream);

            XmlDocument xDoc = new XmlDocument();

            string xml = reader.ReadToEnd();

            xDoc.LoadXml(xml);

            XmlElement xRoot = xDoc.DocumentElement;
            XmlNode pos = xRoot.SelectSingleNode("//pos");


            return pos.InnerXml;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //textBox1.Text = GetCord2(@"https://geocode-maps.yandex.ru/1.x/?apikey=4f98a5df-1522-46aa-8279-2a8691188448&geocode=Кобулети,%20просп.%20Давида%20Агмашенебели,%20291");
        }
    }
}
