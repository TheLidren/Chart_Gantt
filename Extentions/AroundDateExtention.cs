using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chart_Gantt.Extentions
{
    public static class AroundDateExtention
    {
        public static DateTime AroundMethod(this DateTime date)
        {
            date = date.Minute < 15 ? date.AddMinutes(-date.Minute)
                : date.Minute > 15 && date.Minute < 45 ? date.AddMinutes(30 - date.Minute)
                : date.AddMinutes(60 - date.Minute);
            date = date.AddSeconds(-date.Second);
            return date;
        }

    }
}
