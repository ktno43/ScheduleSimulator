using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Test
{
    class Course
    {
        private string course;
        private string courseNum;
        private string courseLoc;
        private string courseDay;
        private string courseStartTime;
        private string courseEndTime;
        private string courseInstr;

        public Course(String[] line)
        {
            course = line[0];
            courseNum = line[1];
            courseLoc = line[2];
            courseDay = line[3];
            courseStartTime = line[4];
            courseEndTime = line[5];
            courseInstr = line[6];
        }

        protected String[] getCourseInfo()
        {
            String[] arrCourseInfo = { this.course,
                this.courseNum,
                this.courseLoc,
                this.courseDay,
                this.courseStartTime,
                this.courseEndTime,
                this.courseInstr };

            return arrCourseInfo;
        }

        protected Boolean isOverlap(Course c1, Course c2)
        {
            if ((!isUnknown(c1) && (!isUnknown(c2))))
            {
                String[] arrC1 = parseTime(c1);
                String[] arrC2 = parseTime(c2);

                if (isDayConflict(c1, c2) && isTimeConflict(arrC1, arrC2))
                {
                    return true;
                }
            }

            return false;
        }

        private Boolean isDayConflict(Course c1, Course c2)
        {
            String c1Day = Regex.Replace(c1.courseDay, ".{2}", "$0,").TrimEnd(',');
            String c2Day = Regex.Replace(c2.courseDay, ".{2}", "$0,").TrimEnd(',');

            String[] arrC1Day = c1Day.Split(',');
            String[] arrC2Day = c2Day.Split(',');

            if (arrC1Day.Length > arrC2Day.Length)
            {
                if (arrC2Day.Intersect(arrC1Day).Any())
                {
                    return true;
                }
            }

            else
            {
                if (arrC1Day.Intersect(arrC2Day).Any())
                {
                    return true;
                }
            }
            return false;
        }


        private Boolean isTimeConflict(String[] t1, String[] t2)
        {
            int t1HrS = Int32.Parse(t1[0]);
            int t1MinS = Int32.Parse(t1[1]);
            DateTime t1TimeS = getTime(t1HrS, t1MinS, t1[3]);

            int t1HrE = Int32.Parse(t1[3]);
            int t1MinE = Int32.Parse(t1[4]);
            DateTime t1TimeE = getTime(t1HrE, t1MinE, t1[5]);


            int t2HrS = Int32.Parse(t2[0]);
            int t2MinS = Int32.Parse(t2[1]);
            DateTime t2TimeS = getTime(t2HrS, t2MinS, t2[3]);

            int t2HrE = Int32.Parse(t2[3]);
            int t2MinE = Int32.Parse(t2[4]);
            DateTime t2TimeE = getTime(t2HrE, t2MinE, t2[5]);

            return (t1TimeS <= t2TimeE) && (t2TimeS <= t1TimeE);
        }


        private DateTime getTime(int hour, int minute, String amPm)
        {
            int year = 2019;
            int month = 9;
            int day = 23;

            if ((amPm.Equals("pm") || amPm.Equals("PM")) && hour < 12)
            {
                hour += 12;
            }

            return new DateTime(year, month, day, hour, minute, 00);
        }

        private Boolean isUnknown(Course c)
        {
            return (c.courseStartTime.Equals("TBA")) &&
                (c.courseDay.Equals("TBA"));
        }

        private String[] parseTime(Course c)
        {
            String[] timeArr = new String[6];

            String sHr = c.courseStartTime.Substring(0, 2);
            String sMin = c.courseStartTime.Substring(3, 2);
            String sAmPm = c.courseStartTime.Substring(5, 2);

            String eHr = c.courseEndTime.Substring(0, 2);
            String eMin = c.courseEndTime.Substring(3, 2);
            String eAmPm = c.courseEndTime.Substring(5, 2);

            timeArr[0] = sHr;
            timeArr[1] = sMin;
            timeArr[2] = sAmPm;

            timeArr[3] = eHr;
            timeArr[4] = eMin;
            timeArr[5] = eAmPm;

            return timeArr;
        }
    }
}
