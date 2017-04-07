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



namespace excel1
{
    class Program
    {
        static object missing = System.Reflection.Missing.Value;

        static void Main(string[] args)
        {
            Excel.Application excel = new Excel.Application();
            string address = "C:\\Users\\Elegant\\Desktop\\1.xls";
            string thanksEnding = "候会，收到烦复！党办小唐";
            string endingPlease = "请您会期关心时间情况通报的群消息，并提前在";
            string beginning = "尊敬的";
            string zidingyu = "兹定于";
            string please = "，请您于";

            try
            {
                excel.Visible = true;
                excel.UserControl = true; //Read-Only Mode;
                Excel.Workbook workbook = excel.Application.Workbooks.Open(address, missing, true, missing, missing,
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
                Excel.Range topics = worksheet.Cells.get_Range("B2", "B" + rowNum);
                Excel.Range presenter = worksheet.Cells.get_Range("C2", "C" + rowNum);
                Excel.Range school_leader = worksheet.Cells.get_Range("D2", "D" + rowNum);
                Excel.Range others = worksheet.Cells.get_Range("E2", "E" + rowNum);
                Excel.Range time = worksheet.Cells.get_Range("F2", "F" + rowNum);

                //From Range to object
                object[,] topic_o = (object[,])topics.Value2;
                object[,] presenter_o = (object[,])presenter.Value2;
                object[,] school_leader_o = (object[,])school_leader.Value2;
                object[,] others_o = (object[,])others.Value2;
                object[,] time_o = (object[,])time.Value2;

                //From Object to String
                string[] topic_str = new string[rowNum - 1];
                string[] presenter_str = new string[rowNum - 1];
                string[] school_leader_str = new string[rowNum - 1];
                string[] others_str = new string[rowNum - 1];
                string[] time_str = new string[rowNum - 1];

                for (int i = 1;i<=rowNum-1;i++)
                {
                    topic_str[i-1] = topic_o[i,1].ToString();
                    presenter_str[i - 1] = presenter_o[i,1].ToString();
                    school_leader_str[i - 1] = school_leader_o[i, 1].ToString();
                    others_str[i - 1] = school_leader_o[i, 1].ToString();
                    time_str[i - 1] = time_o[i, 1].ToString();
                }

                for(int i = 0;i<others_str.Length;i++)
                {

                }

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
