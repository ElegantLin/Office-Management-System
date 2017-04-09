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
            
            string excel_address = "Use your directory";

            string thanksEnding = "收到烦复！党办小唐";
            string endingPlease = "请您会期关心时间情况通报的群消息，";
            string beginning = "尊敬的";
            string zidingyu = "兹定于";
            string please = "，请您于";
            string date = "";
            string place = "";
            Word.Application word = new Word.Application();
            

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

                word.Visible = true;
                Word.Document newdoc;
                newdoc = word.Documents.Add(missing, missing, missing, true);
                newdoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;

                string[,] content = new string[rowNum - 2, colNum - 1];
                
                ArrayList al = new ArrayList();
                string message = ""; // 尊敬的

                for (int i = 3; i <= rowNum - 1; i++)
                {
                    string[] others = worksheet.Cells[i,5].Value.ToString().Split(new char[2] { '、', '，' });
                    string topic = worksheet.Cells[i, 2].Value.ToString();
                    topic = topic.Replace("\n", "");
                    string presenter = (worksheet.Cells[i,3].Value == null) ? "" : worksheet.Cells[i, 3].Value.ToString();
                    string leader = (worksheet.Cells[i,4].Value == null) ? "" : worksheet.Cells[i, 4].Value.ToString();
                    string time = (worksheet.Cells[i,6].Value == null)?"": worksheet.Cells[i, 6].Value.ToString();
                    string am = "";

                    string[] hour = time.Split(new char[3] { '：', '-', ' ' });
                    string[] start = time.Split(new char[2] { '-', ' ' });
                    string startTime = start[0];

                    int hour_i = int.Parse(hour[0]);

                    if (hour_i > 0 && hour_i < 6)
                    {
                        am = "下午";
                    }
                    else
                        am = "上午";

                    foreach (string temp in others)
                    {
                        al.Add(temp);
                    }

                   for(int k = 0;k<al.Count;k++)
                    {
                        message = beginning + al[k].ToString() + '，' + zidingyu + date + am + "在" + place +
                            "召开" + topic + "会议，" + endingPlease + thanksEnding + "\n";
                        newdoc.Paragraphs.Last.Range.Text = message;
                        Console.WriteLine(message);
                    }
                    newdoc.Paragraphs.Last.Range.Text = "\n";
                    message = beginning + leader + '，' + zidingyu + date + am + "在" + place +
                        "召开" + topic + "会议，" + endingPlease + thanksEnding + "\n";
                    newdoc.Paragraphs.Last.Range.Text = message;
                    Console.WriteLine(message);
                    
                }

                word.Quit();
            }
            catch (Exception ex)
            {
                excel.Application.Workbooks.Close();
                excel.Quit();
                Console.WriteLine("Exception" + ex);
            }
            
            excel.Application.Workbooks.Close();
            excel.Quit();

       
        }
    }
}
