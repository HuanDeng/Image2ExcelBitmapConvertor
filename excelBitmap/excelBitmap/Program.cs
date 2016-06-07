using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Drawing;
using System.Collections;
using System.IO;
namespace excelBitmap
{
    class Program
    {
        static byte[] Image2Bytes(Image Pic) 
        {
            MemoryStream ms = new MemoryStream();
            if (Pic == null) return new byte[ms.Length];
            Pic.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            byte[] bPic = new byte[ms.Length];
            bPic = ms.GetBuffer();
            return bPic;
        }
        static void Main(string[] args)
        {
            //ColorFormat cf;
            //cf.

            // read a jpg picture or bmp picture
            // get bitmap Array
            // get ArrayHeight ArrayW
            // copy to excel &save it
            string excelFilePath="temp.xlsx";
            Console.WriteLine("****************************");
            Console.WriteLine("****Bitmap to ExcelBitmap***");
            Console.WriteLine("****                     ***");
            Console.WriteLine("****                     ***");
            Console.WriteLine("****************************");

            //Image bitmapImage=Image.FromFile("temp.bmp");
            //byte[] bitmapArray=Image2Bytes(bitmapImage);

            Bitmap picBitmap = new Bitmap("a.jpg");
            int excelHeight = picBitmap.Height;
            int excelWidth = picBitmap.Width;
            //picBitmap.
            //picBitmap.GetPixel(0, 0);
            Application excelApp = new Application();
            Workbooks excelWbks = excelApp.Workbooks;
            _Workbook excel_Wbk = excelWbks.Add(excelFilePath);
            Sheets excelSheets = excel_Wbk.Sheets;
            _Worksheet excel_Worksheet = (_Worksheet)excelSheets.get_Item(1);
            //excel_Worksheet.Cells[1, 1] = "helloworld";
            //ColorFormat excelColorFormat;

            for (int i = 1; i <= excelHeight; i++) 
            {
                ((Range)excel_Worksheet.Rows[i, Missing.Value]).RowHeight = 4.5;  //4.5=9pix
            }
            for (int i = 1; i <= excelWidth; i++)
            {
                ((Range)excel_Worksheet.Columns[i, Missing.Value]).ColumnWidth = 0.5; //0.5=9pix
            }
            Color picColor;
            ColorConverter colorConv=new ColorConverter();
            int ptr=0;
            for (int i = 0; i < excelHeight; i++) 
            {
                for (int j = 0; j < excelWidth; j++) 
                {
                    picColor = picBitmap.GetPixel(j,i);

                    ((Range)excel_Worksheet.Cells[i + 1, j + 1]).Interior.Color = ((int)picColor.B << 16) + ((int)picColor.G << 8) + ((int)picColor.R);    //RGB
                    Console.WriteLine("x:" + i.ToString() + " " + "y:" + j.ToString());
                }
            }

            //save file
            excelApp.AlertBeforeOverwriting = false;
            excel_Wbk.SaveAs(excelFilePath,Missing.Value,Missing.Value,Missing.Value,Missing.Value
                ,Missing.Value,Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,Missing.Value,Missing.Value
                ,Missing.Value,Missing.Value,Missing.Value);

            excel_Wbk.Close(null,null,null);
            excelWbks.Close();
            excelApp.Quit();
            //Console.ReadKey();

        }
    }
}
