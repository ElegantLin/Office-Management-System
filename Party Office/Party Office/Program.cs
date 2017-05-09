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
using static System.Console;

namespace Party_Office
{
    /// <summary>
    /// One person or place included
    /// </summary>
    class Person
    {
        string name;
        List<int> conference;
        List<bool> presentation;

        public Person(string Name)
        {
            name = Name;
            conference = new List<int>();
            presentation = new List<bool>();
        }

        public string Name { get => name; set => name = value; }

        public void Add(int conferenceNo, bool Presentation)
        {
            conference.Add(conferenceNo);
            presentation.Add(Presentation);
        }
    }

    class Conference
    {
        string title;
        int startHour;
        int startMinute;
        int endHour;
        int endMinute;
        string[] presenter;
        string[] participant;

        public Conference(string Title, string Time, string[] Presentation, string[] Participants)
        {
            title = Title;

            string[] start_end = Time.Trim().Split('-');
            string[] start = start_end[0].Split(new char[2] { ':', '：' });
            string[] end = start_end[1].Split(new char[2] { ':', '：' });
            startHour = int.Parse(start[0]);
            startMinute = int.Parse(start[1]);
            endHour = int.Parse(end[0]);
            endMinute = int.Parse(end[1]);

            presenter = Presentation;
            participant = Participants;
        }

        public string[] Participant { get => participant; set => participant = value; }
        public string[] Presenter { get => presenter; set => presenter = value; }
    }

    class Program
    {
        /// <summary>
        /// Function to get the time
        /// Input:
        /// Return:
        /// </summary>
        /// <returns></returns>
        static string[] GetTimeAddress(Excel.Worksheet sheet)
        {
            string[] TimeAddr = new string[2];
            Excel.Range time_address = sheet.Cells.get_Range("A1");
            string time_address_string = ((object)time_address.Value2).ToString().Trim();
            string[] sArray = time_address_string.Split(new char[2] { ' ', '：' });

            int length = sArray.Length - 1;
            TimeAddr[1] = sArray[length];
            int j = 0;
            for(int i = 0;i<time_address_string.Length;i++)
            { 
                if(time_address_string[i] == ')'|| time_address_string[i] == '）')
                {
                    break;
                }
                else
                {
                    j++;
                }
            }

            string temp = time_address_string.Substring(0, ++j);
            string[] temp1 = temp.Split(new char[2] { ':', '：' });
            TimeAddr[0] = temp1[1];
            return TimeAddr;
        }

        static List<Conference> GetConf(Excel.Worksheet worksheet, int rowNum)
        {
            List<Conference> conf_list = new List<Conference>();
            for (int i = 3; i <= rowNum - 1; i++)
            {
                string topic = worksheet.Cells[i, 2].Value.ToString();
                topic = topic.Replace("\n", "");

                string presenter = (worksheet.Cells[i, 3].Value == null) ? "" : worksheet.Cells[i, 3].Value.ToString();
                string[] presenter_array = presenter.Split(new char[2] { '，', '、' });

                string participant = (worksheet.Cells[i, 5].Value == null) ? "" : worksheet.Cells[i, 5].Value.ToString();
                string[] participant_array = participant.Split(new char[2] { '、', '，' });

                string time = (worksheet.Cells[i, 6].Value == null) ? "" : worksheet.Cells[i, 6].Value.ToString();

                Conference conf = new Conference(topic, time, presenter_array, participant_array);
            }
            return conf_list;
        }

        static List<Person> GetPerson(List<Conference> conf_list)
        {
            List<Person> person_list = new List<Person>();

            int k = 0;
            foreach(Conference conf in conf_list)
            {
                for(int i = 0;i<conf.Participant.Length;i++)
                {
                    if(IsInPersonList(person_list,conf.Participant[i]))
                    {
                        Merge(person_list, conf.Participant[i], k, false);
                    }
                    else
                    {
                        Person per = new Person(conf.Participant[i]);
                        per.Add(k, false);
                        person_list.Add(per);
                    }
                }

                for(int i = 0;i<conf.Presenter.Length;i++)
                {
                    if (IsInPersonList(person_list, conf.Participant[i]))
                    {
                        Merge(person_list, conf.Participant[i], k, true);
                    }
                    else
                    {
                        Person per = new Person(conf.Participant[i]);
                        per.Add(k, true);
                        person_list.Add(per);
                    }
                }
                k++;
            }

            return person_list;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="person_list">Existing Person list</param>
        /// <param name="name">The name to judge</param>
        /// <returns></returns>
        static bool IsInPersonList(List<Person> person_list, string name)
        {
            for(int i =0;i<person_list.Capacity;i++)
            {
                if (person_list[i].Name != name)
                    continue;
                else
                    return true;
            }
            return false;
        }

        static Person Merge(List<Person> person_list, string name, int confNum, bool partOrPre)
        {
            int i = 0;
            for(i = 0;i<person_list.Capacity;i++)
            {
                if(person_list[i].Name == name)
                {
                    break;
                }
            }
            Person per = person_list[i];
            per.Add(confNum, partOrPre);
            return per;
        }

        static void output(List<Person> person_list)
        {
            Word.Application word = new Word.Application();
            word.Visible = true;
            Word.Document newdoc;
            try
            {
                newdoc = word.Documents.Add(missing, missing, missing, true);
                newdoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            }
            catch(Exception e)
            {
                WriteLine(e);
            }
            foreach(Person per in person_list)
            {
                
            }
        }
        

        /// <summary>
        /// Missing Value if null
        /// </summary>
        static object missing = System.Reflection.Missing.Value;

        static void Main(string[] args)
        {
            ///1.Open the program and the sheet.
            Excel.Application excel = new Excel.Application();
            string excel_address = "C:\\Users\\Elegant\\Desktop\\1.xls";
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
                
                string[] TimeAddr = GetTimeAddress(worksheet);

                List<Conference> conf_list = GetConf(worksheet, rowNum);
                List<Person> per_list = GetPerson(conf_list);

            }
            catch (Exception e)
            {
                WriteLine(e);
            }
        }
    }
}
