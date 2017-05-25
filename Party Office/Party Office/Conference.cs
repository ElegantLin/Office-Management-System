using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Party_Office
{
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
            return startHour + ":" + ((startMinute < 10) ? "0" + startMinute.ToString() : startMinute.ToString())
                + "--" + endHour + ":" + ((endMinute < 10) ? "0" + endMinute.ToString() : endMinute.ToString());
        }

        public string[] Participant { get => participant; set => participant = value; }
        public string[] Presenter { get => presenter; set => presenter = value; }
        public string Title { get => title; set => title = value; }
    }
}
