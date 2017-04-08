using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace excel1
{
    class Program
    {
        static object missing = System.Reflection.Missing.Value;

        static void Main(string[] args)
        {
            Excel.Application excel = new Excel.Application();
            string excel_address = "C:\\Users\\Elegant\\Desktop\\1.xls";
            string word_address = "C:\\Users\\Elegant\\Desktop\\1.doc";
            string thanksEnding = "候会，收到烦复！党办小唐";
            string endingPlease = "请您会期关心时间情况通报的群消息，并提前在";
            string beginning = "尊敬的";
            string zidingyu = "兹定于";
            string please = "，请您于";

            try
            {
                excel.Visible = true;
                excel.UserControl = true; //Read-Only Mode;
                Excel.Workbook workbook = excel.Application.Workbooks.Open(excel_address, missing, true, missing, missing,
                    missing, missing, missing, missing, false, missing, missing, missing, missing, missing);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets.get_Item(1);

                //Get the total numbers of rows
                int rowNum = worksheet.UsedRange.Cells.Rows.Count;

                //Get the total numbers of cols
                int colNum = worksheet.UsedRange.Cells.Columns.Count;

                //Ready to write to word
                //Get the date and address
                Excel.Range time_address = worksheet.Cells.get_Range("A1");
                string time_address_string = ((object)time_address.Value2).ToString();
                string[] sArray = time_address_string.Split(new char[2]{' ','：' });
                int j = 0;
                string date, place;
                for(int i = 0;i<sArray.Length;i++)
                {
                    if (sArray[i] == "时间" || sArray[i] == "地点" || sArray[i] == " " || sArray[i] == "")
                        continue;
                    else
                    {
                        if (j != 0)
                        {
                            place = sArray[i];
                            Console.Write("地点是：" + place);
                            break;
                        }
                        date = sArray[i];
                        Console.WriteLine("日期是：" + date);
                        j++;
                    }
                }



                //Get other elements
                //Excel.Range topics = worksheet.Cells.get_Range("B2", "B" + rowNum);
                //Excel.Range presenter = worksheet.Cells.get_Range("C2", "C" + rowNum);
                //Excel.Range school_leader = worksheet.Cells.get_Range("D2", "D" + rowNum);
                //Excel.Range others = worksheet.Cells.get_Range("E2", "E" + rowNum);
                //Excel.Range time = worksheet.Cells.get_Range("F2", "F" + rowNum);


                //From Range to object
                //object[,] topic_o = (object[,])topics.Value2;
                //object[,] presenter_o = (object[,])presenter.Value2;
                //object[,] school_leader_o = (object[,])school_leader.Value2;
                //object[,] others_o = (object[,])others.Value2;
                //object[,] time_o = (object[,])time.Value2;

                Word.Application word = new Word.Application();
                word.Visible = true;
                Word.Document document = word.Documents.Open(word_address, missing, false, true, missing, missing, missing, missing
                    , missing, missing, missing, true, true, missing, missing, missing);
                document.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
                
                string[,] content = new string[rowNum - 2, colNum - 1];
                
                ArrayList al = new ArrayList();
                string message = beginning; // 尊敬的

                for (int i = 2; i <= rowNum - 1; i++)
                {
                    string[] others = worksheet.Cells[i, 4].Value.ToString().Split(new char[2] { '、', '，' });
                    string topic = worksheet.Cells[i, 1].Value.ToString();
                    string presenter = worksheet.Cells[i, 2].Value.ToString();
                    string leader = worksheet.Cells[i, 3].Value.ToString();
                    string time = worksheet.Cells[i, 5].Value.ToString();

                    foreach (string temp in others)
                    {
                        al.Add(temp);
                    }
                    //for (j = 1; j < colNum; j++)
                    //{
                    //    content[i - 2, j - 1] = worksheet.Cells[i, j].Value.ToString();
                    //    if (j == 3)
                    //    {
                    //        string[] others = content[i - 2, j].Split(new char[2] { '、', '，' });
                    //        foreach(string temp in others)
                    //        {
                    //            al.Add(temp);
                    //        }
                    //    }
                    //}
                    foreach (Array array in al)
                    {
                        //message = message + array.ToString() + '，' + zidingyu + content[] 
                    }
                   
                   
                }

                

                //for(int i = 0;i<others_str.Length;i++)
                //{

                //}

            }
            catch (Exception ex)
            {
                
                excel.Application.Workbooks.Close();
                Console.WriteLine("Exception" + ex);
            }
            excel.Application.Workbooks.Close();
       
        }
    }
}
