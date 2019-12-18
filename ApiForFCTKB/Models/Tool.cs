using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApiForFCTKB.Models
{
    public class Tool 
    {
        /// <summary> 
        /// 取指定日期是一年中的第几周 
        /// </summary> 
        /// <param name="dtime">给定的日期</param> 
        /// <returns>数字 一年中的第几周</returns> 
        public static int WeekOfYear(DateTime dtime)
        {
            try
            {
                //确定此时间在一年中的位置
                int dayOfYear = dtime.DayOfYear;
                //当年第一天
                DateTime tempDate = new DateTime(dtime.Year, 1, 1);
                //确定当年第一天
                int tempDayOfWeek = (int)tempDate.DayOfWeek;
                tempDayOfWeek = tempDayOfWeek == 0 ? 7 : tempDayOfWeek;
                ////确定星期几
                int index = (int)dtime.DayOfWeek;
                index = index == 0 ? 7 : index;
                //当前周的范围
                DateTime retStartDay = dtime.AddDays(-(index - 1));
                DateTime retEndDay = dtime.AddDays(6 - index);
                //确定当前是第几周
                int weekIndex = (int)Math.Ceiling(((double)dayOfYear + tempDayOfWeek - 1) / 7);
                if (retStartDay.Year < retEndDay.Year)
                {
                    weekIndex = 1;
                }
                return weekIndex;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public static string GetFormatDateWithDay(DateTime dt)
        {
            string year = dt.Year.ToString().Substring(2);
            string week = WeekOfYear(dt).ToString();
            string day = GetDay(dt).ToString();
            return year + week + "." + day;
        }
        public static int GetDay(DateTime dt)
        {
            int week;
            switch (dt.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    week = 7;
                    break;
                case DayOfWeek.Monday:
                    week = 1;
                    break;
                case DayOfWeek.Tuesday:
                    week = 2;
                    break;
                case DayOfWeek.Wednesday:
                    week = 3;
                    break;
                case DayOfWeek.Thursday:
                    week = 4;
                    break;
                case DayOfWeek.Friday:
                    week = 5;
                    break;
                case DayOfWeek.Saturday:
                    week = 6;
                    break;
                default:
                    week = 0;
                    break;
            }
            return week;
        }
    }
}