using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DBProject
{
    class Teacher
    {
        private static int Counter = 0;

        public int Id;
        String Position;
        String Initials;
        String Lastname;
        public String Name;

        public Teacher(String Name)
        {
            this.Id = Counter++;
            this.Name = Name;
            string[] output = Name.Split(' ');
            this.Position = "lol";
            this.Initials = "kek";
            this.Lastname = Name;

        }

        override public String ToString()
        {
            return "('" + this.Id + "', '" + this.Lastname + "', '" + this.Initials + "', '" + this.Position + "')";
        }
    }
}
