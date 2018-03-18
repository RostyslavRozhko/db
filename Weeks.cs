using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DBProject
{
    class Weeks
    {
        public int EntityId;
        public List<String> WeeksList;

        public Weeks(int EntityId, List<String> WeeksList)
        {
            this.EntityId = EntityId;
            this.WeeksList = WeeksList;
        }
    }
}
