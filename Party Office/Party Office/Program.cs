using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using static System.Console;

namespace Party_Office
{
    /// <summary>
    /// One person or place included
    /// </summary>
    class Person
    {
        string name;
        List<int> presentation;
        List<int> participant;

        public Person(string Name)
        {
            name = Name;
            presentation = new List<int>();
            participant = new List<int>();
        }

        public string Name { get => name; set => name = value; }
        public List<int> Presentation { get => presentation; set => presentation = value; }
        public List<int> Participant { get => participant; set => participant = value; }

        public void AddPre(int preNo)
        {
            presentation.Add(preNo);
        }

        public void AddPar(int ParNo)
        {
            participant.Add(ParNo);
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

        public string Time()
        {
            return startHour + ":" + startMinute + "--" + endHour +":" + endMinute;
        }

        public string[] Participant { get => participant; set => participant = value; }
        public string[] Presenter { get => presenter; set => presenter = value; }
        public string Title { get => title; set => title = value; }
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
                conf_list.Add(conf);
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
                        ///Not Presentation
                        person_list = Merge(person_list, conf.Participant[i], k, false);
                    }
                    else
                    {
                        Person per = new Person(conf.Participant[i]);
                        ///
                        per.AddPar(k);
                        person_list.Add(per);
                    }
                }

                for(int i = 0;i<conf.Presenter.Length;i++)
                {
                    if (IsInPersonList(person_list, conf.Presenter[i]))
                    {
                        ///Presentation
                        person_list = Merge(person_list, conf.Presenter[i], k, true);
                    }
                    else
                    {
                        Person per = new Person(conf.Presenter[i]);
                        per.AddPre(k);
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
            for(int i =0;i<person_list.Count;i++)
            {
                if (person_list[i].Name != name)
                    continue;
                else
                    return true;
            }
            return false;
        }

        static List<Person> Merge(List<Person> person_list, string name, int confNum, bool partOrPre)
        {
            int i = 0;
            for(i = 0;i<person_list.Count;i++)
            {
                if(person_list[i].Name == name)
                {
                    if (partOrPre)
                        person_list[i].AddPre(confNum);
                    else
                        person_list[i].AddPar(confNum);
                    break;
                }
            }
            return person_list;
        }

        static void output(List<Person> person_list, string[] TimeAdd, List<Conference> conf_list)
        {
            Word.Application word = new Word.Application();
            word.Visible = true;
            Word.Document newdoc;
            newdoc = word.Documents.Add(missing, missing, missing, true);
            newdoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            foreach (Person per in person_list)
            {
                try
                {
                    string str1 = "尊敬的" + per.Name + "兹定于" + TimeAdd[0] + "在" + TimeAdd[1] + "召开党委常委会，请您于" + '\n';
                    newdoc.Paragraphs.Last.Range.Text = str1;
                    //newdoc.Paragraphs.Last.Range.Text = "\n";
                    //myPag.Range.ListFormat.ApplyBulletDefault();

                    //Participant 
                    for (int i = 0;i<per.Participant.Count;i++)
                    {
                        int j = per.Participant[i];
                        Conference con = conf_list[j];
                        string subStr = con.Time() + "列席第" + (j + 1).ToString() + "个议题" + (j + 1).ToString() + "." + con.Title + "\n";
                        newdoc.Paragraphs.Last.Range.Text = subStr;
                    }

                    //Present 
                    for(int i = 0;i<per.Presentation.Count;i++)
                    {
                        int j = per.Presentation[i];
                        Conference con = conf_list[i];
                        string subStr = con.Time() + "汇报第" + (j + 1).ToString() + "个议题" + (j+1).ToString() + "." + con.Title + "\n";
                        newdoc.Paragraphs.Last.Range.Text = subStr;
                    }

                    string subStr1 = "请您会期关心时间情况通报的群消息，并提前到会，收到烦复！" + "\n" + "党办小唐";
                    newdoc.Paragraphs.Last.Range.Text = subStr1;
                    
                }
                catch(Exception e)
                {
                    WriteLine(e);
                }
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
                output(per_list, TimeAddr, conf_list);

            }
            catch (Exception e)
            {
                WriteLine(e);
            }
        }
    }
}
