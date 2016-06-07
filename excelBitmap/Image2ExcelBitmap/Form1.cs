using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Collections;
using System.IO;
using System.Threading;

namespace Image2ExcelBitmap
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Bitmap imageBitmap;
        string excelFilePath;
        private void StartButton_Click(object sender, EventArgs e)
        {
            if (StartButton.Text == "Open Image File") 
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "(*.jpg)|*.jpg";
                if (DialogResult.OK == ofd.ShowDialog()) 
                {
                    try
                    {
                        imageBitmap = new Bitmap(ofd.FileName);
                        ViewPicureBox.Image = Image.FromFile(ofd.FileName);
                        ImageFilePathTextBox.Text = ofd.FileName;
                        ImageInfoRichTextBox.AppendText(imageBitmap.Size.ToString());
                        StartButton.Text = "Save Excel File";
                    }
                    catch (Exception ex) 
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }
            else if (StartButton.Text == "Save Excel File") 
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "(*.xlsx)|*.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK) 
                {
                    excelFilePath = sfd.FileName;
                    ExcelFilePathTextBox.Text = excelFilePath;
                    StartButton.Text = "Convert";
                }
            }
            else if (StartButton.Text == "Convert")
            {
                Thread thd = new Thread(new ThreadStart( Image2Excel));
                thd.Start();
                StartButton.Enabled = false;
            }
        }
        void Image2Excel() 
        {
            Bitmap bmp = imageBitmap;
            int excelHeight = bmp.Height;
            int excelWidth = bmp.Width;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbooks excelWbks = excelApp.Workbooks;
            _Workbook excel_Wbk = excelWbks.Add(true);
            Sheets excelSheets = excel_Wbk.Sheets;
            _Worksheet excel_Worksheet = (_Worksheet)excelSheets.get_Item(1);

            for (int i = 1; i <= excelHeight; i++)
            {
                ((Range)excel_Worksheet.Rows[i, Missing.Value]).RowHeight = 4.5;  //4.5=9pix
            }
            for (int i = 1; i <= excelWidth; i++)
            {
                ((Range)excel_Worksheet.Columns[i, Missing.Value]).ColumnWidth = 0.5; //0.5=9pix
            }
            Color picColor;
            ProcessProgressBar.Minimum = 0;
            ProcessProgressBar.Maximum = excelHeight-1;

            for (int i = 0; i < excelHeight; i++)
            {
                for (int j = 0; j < excelWidth; j++)
                {
                    picColor = bmp.GetPixel(j, i);

                    ((Range)excel_Worksheet.Cells[i + 1, j + 1]).Interior.Color = ((int)picColor.B << 16) + ((int)picColor.G << 8) + ((int)picColor.R);    //RGB
                    //Console.WriteLine("x:" + i.ToString() + " " + "y:" + j.ToString());
                }
                ProcessProgressBar.Value = i;
                ProcessLable.Text = i.ToString() + "/" + (excelHeight-1).ToString();
            }
            //save file
            excelApp.AlertBeforeOverwriting = false;
            excel_Wbk.SaveAs(excelFilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value
                , Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value
                , Missing.Value, Missing.Value, Missing.Value);

            excel_Wbk.Close(null, null, null);
            excelWbks.Close();
            excelApp.Quit();
            MessageBox.Show("Convert Successful");
            StartButton.Enabled = true;
            StartButton.Text = "Open Image File";
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
