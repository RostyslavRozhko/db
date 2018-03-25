using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DBProject
{
    class ExcelRecord
    {
        static int counter = 0;

        String Year;
        String Speciality;
        String DayOfWeek;
        String Time;
        String Subject;
        String Teacher;
        String Group;
        String Room;
        String Weeks;
        public int Id;

        public ExcelRecord(String Year, String Speciality, String DayOfWeek, String Time, String Subject, String Teacher, String Group, String Room, String Weeks)
        {
            this.Id = counter++;
            this.Year = Year;
            this.Speciality = Speciality;
            this.DayOfWeek = DayOfWeek;
            this.Time = Time;
            this.Subject = Subject;
            this.Teacher = Teacher;
            this.Group = Group;
            this.Room = Room;
            this.Weeks = Weeks;
        }

        public override string ToString()
        {
            return "'" + this.Speciality + "', " + this.Year + ", " +  this.Id + ", " + this.DayOfWeek + ", " + this.Time + ", '" + this.Room + "', '" + this.Subject + "', '" + this.Group + "', " + this.Teacher + ", '" + this.Weeks + "'";
        }
    }
}
