using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System;

namespace ExcelDosyaOkuma
{
    class Program
    {
        
        static void Main(string[] args)
        {
            
            //string url = "https://flagsapi.com/"+a+"/shiny/64.png";
            //string dosyaAdi = "Value2.png";

            Excel.Application excelApp = new Excel.Application();


            Excel.Workbook workbook = excelApp.Workbooks.Open(@"C:\Kitap.xlsx");


            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];


            for (int row = 1; row <= worksheet.UsedRange.Rows.Count; row++)
            {
                string a = worksheet.Cells[row, 1].Value2;
                if(string.IsNullOrWhiteSpace(a) )
                {
                    continue;
                }
                //Console.Write(worksheet.Cells[row, 1].Value2 + "\t");

                string url = "https://flagsapi.com/"+ a +"/shiny/64.png";
                string dosyaAdi = a + ".png";
                WebClient webClient = new WebClient();

                try
                {
                    webClient.DownloadFile(url, dosyaAdi);
                    Console.WriteLine("Dosya indirildi: " + dosyaAdi);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Hata oluştu: " + ex.Message);
                }

                Console.WriteLine();
            }
            //excell nesnleri temizliyor
            workbook.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


        }
    }
}





