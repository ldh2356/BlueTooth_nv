using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SSES_Program
{
    public class TimeConverter
    {
        public static long ConvertTick32ToTick64(int tick32)
        {
            DateTime dt70 = new DateTime(1970, 1, 1, 9, 0, 0, 0);
            long tick64 = 0;
            tick64 = tick32;
            tick64 *= 10000000L;
            tick64 += dt70.Ticks;

            return tick64;
        }

        public static string ConvertTick32ToTimeString(int tick32)
        {
            long tick64 = ConvertTick32ToTick64(tick32);
            DateTime dt = new DateTime(tick64);
            return DateTimeToString(dt);
        }

        public static int ConvertTick64ToTick32(long tick64)
        {
            DateTime dt70 = new DateTime(1970, 1, 1, 9, 0, 0, 0);
            int tick32 = 0;
            tick64 -= dt70.Ticks;
            tick64 /= 10000000L;
            tick32 = (int)tick64;

            return tick32;
        }

        public static string DateTimeToString(DateTime dateTime)
        {
            string time = string.Format("{0,4:0000}-{1,2:00}-{2,2:00} {3,2:00}:{4,2:00}:{5,2:00}"
                        , dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour, dateTime.Minute, dateTime.Second);
            return time;
        }

        public static DateTime StringToDateTime(string time)
        {
            int year = int.Parse(time.Substring(0, 4));
            int mon = int.Parse(time.Substring(5, 2));
            int date = int.Parse(time.Substring(8, 2));
            int hour = int.Parse(time.Substring(11, 2));
            int min = int.Parse(time.Substring(14, 2));
            int sec = int.Parse(time.Substring(17, 2));

            DateTime dateTime = new DateTime(year, mon, date, hour, min, sec, 0);
            return dateTime;
        }

        /*
        public static DateTime StringEngToDateTime(string time)
        {
            int mon = int.Parse(time.Substring(0, 2));
            int date = int.Parse(time.Substring(3, 2));
            int year = int.Parse(time.Substring(6, 4));
            int hour = int.Parse(time.Substring(11, 2));
            int min = int.Parse(time.Substring(14, 2));
            int sec = int.Parse(time.Substring(17, 2));

            DateTime dateTime = new DateTime(year, mon, date, hour, min, sec, 0);
            return dateTime;
        }
        */
    }
}
