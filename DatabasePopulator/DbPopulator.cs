using System;
using System.IO;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace DbPopulator
{
    [TestClass]
    public class DbPopulator
    {
        private static ChromeDriver driver = null;
        private static StreamWriter swLog = null;
        private static StreamWriter swParseLog = null;

        const int SECOND = 1000;
        const int MILLISECOND = 1;


        [ClassInitialize]
        public static void initialize(TestContext context)
        {
            driver = new ChromeDriver(); // Initialize chrome driver
            driver.Navigate().GoToUrl("https://mynorthridge.csun.edu/psp/PANRPRD/EMPLOYEE/SA/c/NR_SSS_COMMON_MENU.NR_SSS_SOC_BASIC_C.GBL?"); // Navigate to URL
            driver.Manage().Window.Maximize(); // Maximize window
            swLog = new StreamWriter("log.txt"); // Log.txt
            swParseLog = new StreamWriter("log_parse.txt"); // log_parse.txt
        }

        [ClassCleanup]
        public static void cleanup()
        {
            driver.Close(); // Close current browser window currently in focus

            Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver"); // Kill all chrome drivers

            foreach (var chromeDriverProcess in chromeDriverProcesses)
            {
                chromeDriverProcess.Kill();
            }

            swLog.Close(); // Close the log
            swParseLog.Close(); // close the parse log
        }

        private Boolean isWheelGone()
        {
            for (int i = 0; i < 300; i++) // Wait for a maximum of 30 seconds (300 * 100ms = 30,000 ms = 30 second)
            {
                Thread.Sleep(100 * MILLISECOND); // Wait for 1 ms

                IWebElement wheel = driver.FindElementById("processing"); // Find processing wheel

                if (!isElementVisible(wheel)) // If wheel is NOT visible
                {
                    return true; // wheel is gone
                }
            }

            return false; // Problem at this point
        }

        private Boolean isElementVisible(IWebElement element)
        {
            return element.Displayed && element.Enabled; // Return status of element being displayed and enabled
        }

        private void fixIllegalChar(ref String refStr)
        {
            String invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars()); // Get a string of invalid characters

            foreach (char c in invalidChars)
            {
                refStr = refStr.Replace(c.ToString(), ""); // Remove invalid characters
            }
        }

        public static void waitForJSLoad(int timeoutSec = 15)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSec));
            wait.Until(wd => js.ExecuteScript("return document.readyState").ToString() == "complete");
        }

        private String getInnerText(HtmlDocument doc, String id)
        {
            return doc.GetElementbyId(id).InnerText; // Get the text enclosed in the element
        }

        private String removeNewline(String str)
        {
            return str.Replace("\r\n", string.Empty); // Remove the newline in a string
        }

        private void fixCsvChar(ref String refData)
        {
            if (refData.Contains("\"")) // Fix double quotes
            {
                refData = refData.Replace("\"", "\"\"");
            }

            if (refData.Contains(","))  // Fix commas
            {
                refData = String.Format("\"{0}\"", refData);
            }

            if (refData.Contains(System.Environment.NewLine)) // Fix new lines
            {
                refData = String.Format("\"{0}\"", refData);
            }
        }

        [TestMethod]
        public void scrapCourseDb()
        {
            driver.SwitchTo().Frame("ptifrmtgtframe"); // Switch to content frame
            SelectElement subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list

            IList<IWebElement> subjectList = subjectDdList.Options; //Get list of IWeb elements
            int numSubjects = subjectList.Count; // Number of subjects CSUN has

            int numSections = 0; // Number of sections

            // Course
            String courseHtml = "";
            String courseID = "SOC_DETAIL$"; // Prefix ID for courses
            int numCourses = 0;

            // Course section 
            String courseSecHtml = "";
            String courseSecID = "NR_SSS_SOC_NWRK_DESCR15$";

            // Course number
            String courseNumHtml = "";
            String courseNum = "";
            String courseNumID = "win0divNR_SSS_SOC_NSEC_CLASS_NBR$";

            // Course location
            String courseLocHtml = "";
            String courseLoc = "";
            String courseLocID = "win0divMAP$";

            // Course days
            String courseDayHtml = "";
            string courseDay = "";
            String courseDayID = "win0divNR_SSS_SOC_NWRK_DESCR20$";

            // Course times
            String courseTimeHtml = "";
            String courseTime = "";
            String courseStartTime = "";
            String courseEndTime = "";
            String courseTimeID = "win0divNR_SSS_SOC_NSEC_DESCR25_2$";

            // Course instructor
            String courseInstrHtml = "";
            String courseInstr = "";
            String courseInstrID = "win0divFACURL$";

            // Course description
            String courseDescrHtml = "";
            String courseDescr = "";
            String courseDescrID = "NR_SSS_SOC_NWRK_DESCR100_2$";
            String courseTitle = "";

            String subjName = "";
            String subjTitle = "";

            for (int subjectIndex = 1; subjectIndex < numSubjects; subjectIndex++)
            {
                subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list
                subjectDdList.SelectByIndex(subjectIndex); // Select subject

                subjName = driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT").GetAttribute("value"); // Get the value of the selected item in the drop down list

                fixIllegalChar(ref subjName); // Fix any illegal characters

                Thread.Sleep(1 * SECOND); // Wait to click the search button
                driver.FindElementById("NR_SSS_SOC_NWRK_BASIC_SEARCH_PB").Click(); // Click the search button
                Thread.Sleep(1 * SECOND); // Wait for number of sections

                courseHtml = driver.FindElementByClassName("PSGRIDCOUNTER").Text; // String representation of number of sections offered in that course
                numCourses = Int32.Parse(courseHtml.Substring(courseHtml.LastIndexOf(' ') + 1)); // Get number of courses under specified subject

                subjTitle = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")).SelectedOption.Text;

                using (var wCsv = new StreamWriter(subjName + "_PARSE.csv"))
                {
                    wCsv.WriteLine(String.Format("{0},{1},{2},{3},{4},{5},{6}", "Course", "Number", "Location", "Days", "Start Time", "End Time", "Instructor")); // headers
                    wCsv.Flush();

                    swParseLog.Write(DateTime.Now); // parse log subject start
                    swParseLog.WriteLine("   " + subjTitle + ": " + numCourses + " course(s)"); // parse log subject name

                    for (int courseIndex = 0, courseRow = 0; courseIndex < numCourses; courseIndex++) // For the number of courses under the specified subject
                    {
                        if (isWheelGone()) // If the processing wheel is gone, click!
                        {
                            driver.FindElementById(courseID + courseIndex).Click(); // Expand every course through click command

                            if (isWheelGone())
                            {
                                courseSecHtml = driver.FindElementById(courseSecID + courseIndex).Text; // Get the string of offered sections of that course
                                numSections = Int32.Parse(Regex.Match(courseSecHtml, @"\d+").Value); // Convert that string to a number

                                courseDescrHtml = driver.FindElementById(courseDescrID + courseIndex).Text;
                                courseTitle = courseDescrHtml.Substring(0, courseDescrHtml.IndexOf('(')).Trim();
                                courseDescr = courseDescrHtml.Substring(0, courseDescrHtml.IndexOf('-')).Trim(); // Trim the description to show only the class and section

                                swParseLog.WriteLine(courseTitle + ": " + numSections + " section(s)");

                                for (int sectionIndex = 0; sectionIndex < numSections; sectionIndex++, courseRow++) // For the number of sections within the class
                                {
                                    courseNumHtml = driver.FindElementById(courseNumID + courseRow).Text; // Get the course number 
                                    courseNum = removeNewline(courseNumHtml);

                                    courseLocHtml = driver.FindElementById(courseLocID + courseRow).Text; // Get the location of the course
                                    courseLoc = removeNewline(courseLocHtml);

                                    courseDayHtml = driver.FindElementById(courseDayID + courseRow).Text; // Get the instruction days the class is taught
                                    courseDay = removeNewline(courseDayHtml);
                                    if (courseDay.Equals("&nbsp;")) // If day is empty, the string is TBA
                                        courseDay = "TBA";

                                    courseTimeHtml = driver.FindElementById(courseTimeID + courseRow).Text; // Get the time the class is taught
                                    courseTime = removeNewline(courseTimeHtml);

                                    if (courseTime.Contains("-")) // If the class has a time
                                    {
                                        courseStartTime = courseTime.Substring(0, courseTime.IndexOf('-'));
                                        courseEndTime = courseTime.Substring(courseTime.LastIndexOf('-') + 1);
                                    }

                                    courseInstrHtml = driver.FindElementById(courseInstrID + courseRow).Text; // Get the name of the instructor who is teaching the class
                                    courseInstr = removeNewline(courseInstrHtml);

                                    // Fix for .csv format
                                    fixCsvChar(ref courseDescr);
                                    fixCsvChar(ref courseNum);
                                    fixCsvChar(ref courseLoc);
                                    fixCsvChar(ref courseDay);
                                    fixCsvChar(ref courseStartTime);
                                    fixCsvChar(ref courseEndTime);
                                    fixCsvChar(ref courseInstr);

                                    // write to the stream writer with information above
                                    var row = String.Format("{0},{1},{2},{3},{4},{5},{6}",
                                        courseDescr, courseNum, courseLoc, courseDay, courseStartTime, courseEndTime, courseInstr);

                                    wCsv.WriteLine(row);
                                    wCsv.Flush();
                                    swParseLog.WriteLine("Section " + (sectionIndex + 1) + " of " + numSections + ": " + courseNum + " parsed successfully."); // parse log
                                }
                                swParseLog.WriteLine();
                            }

                            else // Problem if here
                            {
                                swParseLog.WriteLine(DateTime.Now);
                                swParseLog.WriteLine("ERROR IN " + subjTitle + ": Problem occurred while trying to parse section.");
                            }
                        }

                        else // Problem if here
                        {
                            swLog.WriteLine(DateTime.Now);
                            swLog.WriteLine("ERROR IN " + subjTitle + ": Problem occurred while trying to parse section.");
                        }
                    }
                }
                swParseLog.WriteLine("\r\n\r\n");

                swLog.Write(DateTime.Now);
                swLog.WriteLine("   " + "Finished parsing subject: " + subjTitle + "\r\n\r\n");

                driver.Navigate().Refresh(); // Refresh page and switch course
                driver.SwitchTo().Frame("ptifrmtgtframe"); // Switch to content frame
            }
        }
    }
}