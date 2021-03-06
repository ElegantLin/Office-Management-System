﻿using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using static System.Console;
using System.Runtime.InteropServices;

namespace Party_Office
{
    /// <summary>
    /// One person or place included
    /// </summary>

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
            for (int i = 0; i < time_address_string.Length; i++)
            {
                if (time_address_string[i] == ')' || time_address_string[i] == '）')
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
            for (int i = 7; i <= rowNum; i++)
            {
                string topic = worksheet.Cells[i, 2].Value.ToString();
                topic = topic.Replace("\n", "");

                string presenter = (worksheet.Cells[i, 3].Value == null) ? "" : worksheet.Cells[i, 3].Value.ToString();
                string[] presenter_array = presenter.Split(new char[2] { '，', '、' });

                string participant = (worksheet.Cells[i, 5].Value == null) ? "" : worksheet.Cells[i, 5].Value.ToString();
                string[] participant_array = participant.Split(new char[2] { '、', '，' });

                string time = (worksheet.Cells[i, 6].Value == null) ? "" : worksheet.Cells[i, 6].Value.ToString();

                string index = (worksheet.Cells[i, 1].Value == null) ? "" : worksheet.Cells[i, 1].Value.ToString();
                Conference conf = new Conference(topic, time, presenter_array, participant_array, index);
                conf_list.Add(conf);
            }
            return conf_list;
        }

        static List<Person> GetPerson(List<Conference> conf_list)
        {
            List<Person> person_list = new List<Person>();

            int k = 0;
            foreach (Conference conf in conf_list)
            {
                for (int i = 0; i < conf.Presenter.Length; i++)
                {

                    if (IsInPersonList(person_list, conf.Presenter[i]))
                    {
                        person_list = Merge(person_list, conf.Presenter[i], k, true);
                    }
                    else
                    {
                        person_list = AddPresenter(person_list, conf.Presenter[i], k);
                    }
                }

                for (int i = 0; i < conf.Participant.Length; i++)
                {
                    if (IsInPersonList(person_list, conf.Participant[i]))
                    {
                        person_list = Merge(person_list, conf.Participant[i], k, false);
                    }
                    else
                    {
                        person_list = AddParticipant(person_list, conf.Participant[i], k);
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
            for (int i = 0; i < person_list.Count; i++)
            {
                if (person_list[i].Name != name)
                    continue;
                else
                    return true;
            }
            return false;
        }

        static List<Person> Merge(List<Person> person_list, string name, int confNum, bool preOrNot)
        {
            for (int i = 0; i < person_list.Count; i++)
            {
                if (person_list[i].Name == name)
                {
                    person_list[i].AddConf(confNum, preOrNot);
                }
                else
                    continue;
            }
            return person_list;
        }

        static List<Person> AddPresenter(List<Person> person_list, string name, int confNum)
        {
            Person per = new Person(name);
            per.AddConf(confNum, true);
            person_list.Add(per);
            return person_list;
        }

        static List<Person> AddParticipant(List<Person> person_list, string name, int confNum)
        {
            Person per = new Person(name);
            per.AddConf(confNum, false);
            person_list.Add(per);
            return person_list;
        }

        static void quitExcel(Excel.Application excel, Excel.Workbook workbook)
        {
            excel.ActiveWorkbook.Close(false, missing, missing);
            excel.Quit();
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excel);
            excel = null;
        }

        static void output(List<Person> person_list, List<Conference> conf_list, Excel.Worksheet worksheet)
        {
            Word.Application word = new Word.Application
            {
                Visible = true
            };
            Word.Document newdoc;
            newdoc = word.Documents.Add(missing, missing, missing, true);
            newdoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;

            object unite = Word.WdUnits.wdStory;


            char symbol = (char)(9632);
            foreach (Person per in person_list)
            {
                try
                {
                    //word.Selection.EndKey(ref unite, ref missing);
                    string str = worksheet.Cells[1, 1].Value.ToString() + per.Name + worksheet.Cells[1, 2].Value.ToString()
                        + worksheet.Cells[2, 1].Value.ToString() + worksheet.Cells[2, 2].Value.ToString() + worksheet.Cells[2, 3].Value.ToString()
                        + worksheet.Cells[2, 4].Value.ToString() + worksheet.Cells[2, 5].Value.ToString() + "\n";
                    word.ActiveDocument.Characters.Last.Select();  // Line 1
                    word.Selection.Collapse();                     // Line 2
                    word.Selection.TypeText(str);
                    //newdoc.Paragraphs.Last.Range.Text = "\n";
                    //myPag.Range.ListFormat.ApplyBulletDefault();
                    for (int i = 0; i < per.Confer.Count; i++)
                    {

                        if (per.PreOrNot[i])
                        {
                            //word.Selection.EndKey(ref unite, ref missing);
                            string subStr = symbol.ToString() + conf_list[per.Confer[i]].Time() + "汇报第" + (per.Confer[i] + 1).ToString() + "个议题"
                                + (per.Confer[i] + 1).ToString() + "." + conf_list[per.Confer[i]].Title + "\n";
                            //word.Selection.EndKey(ref unite, ref missing);
                            word.ActiveDocument.Characters.Last.Select();  // Line 1
                            word.Selection.Collapse();                     // Line 2
                            word.Selection.TypeText(subStr);
                        }
                        else
                        {
                            string subStr = symbol.ToString() + conf_list[per.Confer[i]].Time() + "列席第" + (per.Confer[i] + 1).ToString() + "个议题"
                                + (per.Confer[i] + 1).ToString() + "." + conf_list[per.Confer[i]].Title + "\n";
                            //word.Selection.EndKey(ref unite, ref missing);
                            word.ActiveDocument.Characters.Last.Select();  // Line 1
                            word.Selection.Collapse();                     // Line 2
                            word.Selection.TypeText(subStr);
                        }

                    }


                    word.ActiveDocument.Characters.Last.Select();  // Line 1
                    word.Selection.Collapse();                     // Line 2
                    word.Selection.TypeText(worksheet.Cells[4, 1].Value.ToString() + "\n" + "\n");
                }

                catch (Exception e)
                {
                    WriteLine(e);
                }

                //word.Selection.EndKey(ref unite, ref missing);
                //newdoc.Paragraphs.Last.Range.Text = "\n";
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
            string excel_address = "C:\\3.xls";
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

                List<Conference> conf_list = GetConf(worksheet, rowNum);

                List<Person> per_list = GetPerson(conf_list);

                output(per_list, conf_list, worksheet);

                quitExcel(excel, workbook);

            }
            catch (Exception e)
            {
                WriteLine(e);
            }
        }
    }
}
