using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using HtmlAgilityPack;
using System.IO;
using System.Text.RegularExpressions;

namespace DbPopulator
{
    [TestClass]
    public class CourseParser
    {
        [TestMethod]
        public void comSciParser()
        {
            String courseSecHtml = "";
            String courseSecID = "NR_SSS_SOC_NWRK_DESCR15$";
            int numSections = 0;

            String courseNumHtml = "";
            String courseNum = "";
            String courseNumID = "win0divNR_SSS_SOC_NSEC_CLASS_NBR$";

            String courseLocHtml = "";
            String courseLoc = "";
            String courseLocID = "win0divMAP$";

            String courseDayHtml = "";
            string courseDay = "";
            String courseDayID = "win0divNR_SSS_SOC_NWRK_DESCR20$";

            String courseTimeHtml = "";
            String courseTime = "";
            String courseTimeID = "win0divNR_SSS_SOC_NSEC_DESCR25_2$";

            String courseInstrHtml = "";
            String courseInstr = "";
            String courseInstrID = "win0divFACURL$";

            String courseDescrHtml = "";
            String courseDescr = "";
            String courseDescrID = "NR_SSS_SOC_NWRK_DESCR100_2$";

            HtmlDocument doc = new HtmlDocument();
            String filePath = @"F:\Projects\aspNet\DbPopulator\bin\Debug\COMP_DOM.txt";
            doc.Load(filePath);

            String coursesHtml = doc.GetElementbyId("PSCENTER").InnerText;
            int numCourses = Int32.Parse(coursesHtml.Substring(coursesHtml.LastIndexOf(' ') + 1));

            StreamWriter sw = new StreamWriter("COMP_parse.txt");

            for (int courseIndex = 0, courseRow = 0; courseIndex < numCourses; courseIndex++)
            {
                courseDescrHtml = doc.GetElementbyId(courseDescrID + courseIndex).InnerText;
                courseDescr = courseDescrHtml.Substring(0, courseDescrHtml.IndexOf('-')).Trim();


                courseSecHtml = doc.GetElementbyId(courseSecID + courseIndex).InnerText;
                numSections = Int32.Parse(Regex.Match(courseSecHtml, @"\d+").Value);

                //sw.WriteLine(courseDescrHtml + ": " + numSections + " section(s)");
                for (int sectionIndex = 0; sectionIndex < numSections; sectionIndex++, courseRow++)
                {
                    courseNumHtml = doc.GetElementbyId(courseNumID + courseRow).InnerText;
                    courseNum = courseNumHtml.Replace("\r\n", string.Empty);

                    courseLocHtml = doc.GetElementbyId(courseLocID + courseRow).InnerText;
                    courseLoc = courseLocHtml.Replace("\r\n", string.Empty);

                    courseDayHtml = doc.GetElementbyId(courseDayID + courseRow).InnerText;
                    courseDay = courseDayHtml.Replace("\r\n", string.Empty);
                    if (courseDay.Equals("&nbsp;"))
                        courseDay = "TBA";

                    courseTimeHtml = doc.GetElementbyId(courseTimeID + courseRow).InnerText;
                    courseTime = courseTimeHtml.Replace("\r\n", string.Empty);

                    courseInstrHtml = doc.GetElementbyId(courseInstrID + courseRow).InnerText;
                    courseInstr = courseInstrHtml.Replace("\r\n", string.Empty);

                    sw.WriteLine(courseDescr);
                    sw.WriteLine(courseNum);
                    sw.WriteLine(courseLoc);
                    sw.WriteLine(courseDay);
                    sw.WriteLine(courseTime);
                    sw.WriteLine(courseInstr + "\r\n");
                }
                //    sw.WriteLine("\r\n");
            }
            sw.Close();
        }
    }
}
