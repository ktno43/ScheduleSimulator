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
        private static StreamWriter LOG = null;

        const int SECOND = 1000;
        const int MILLISECOND = 1;


        [ClassInitialize]
        public static void initialize(TestContext context)
        {
            driver = new ChromeDriver(); // Initialize chrome driver
            driver.Navigate().GoToUrl("https://mynorthridge.csun.edu/psp/PANRPRD/EMPLOYEE/SA/c/NR_SSS_COMMON_MENU.NR_SSS_SOC_BASIC_C.GBL?"); // Navigate to URL
            driver.Manage().Window.Maximize(); // Maximize window
            LOG = new StreamWriter("log.txt"); // Log.txt
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

            LOG.Close(); // Close the log
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
            return doc.GetElementbyId(id).InnerText;
        }

        private String removeNewline(String str)
        {
            return str.Replace("\r\n", string.Empty);
        }

        [TestMethod]
        public void scrapCourseDb()
        {
            driver.SwitchTo().Frame("ptifrmtgtframe"); // Switch to content frame
            SelectElement subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list

            IList<IWebElement> subjectList = subjectDdList.Options; //Get list of IWeb elements
            int numSubjects = subjectList.Count; // Number of subjects CSUN has

            String courseID = "SOC_DETAIL$"; // Prefix ID for courses
            Boolean bMoreCourses = true; // Boolean for while loop
            int section = 0; // Section index

            String strSections = ""; // String representation of number of sections
            int numSections = 0; // Number of sections

            String courseName = "";

            for (int j = 30; j < 31; j++)
            {
                section = 0; // Section starts at index 0
                bMoreCourses = true; // Always start with true since they're always more courses at this point

                subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list
                subjectDdList.SelectByIndex(j); // Select major

                courseName = driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT").GetAttribute("value");

                fixIllegalChar(ref courseName);


                Thread.Sleep(1 * SECOND); // Wait to click the search button
                driver.FindElementById("NR_SSS_SOC_NWRK_BASIC_SEARCH_PB").Click(); // Click the search button
                Thread.Sleep(1 * SECOND + 500 * MILLISECOND); // Wait for number of sections


                strSections = driver.FindElementByClassName("PSGRIDCOUNTER").Text; // String representation of number of sections offered in that course

                numSections = Int32.Parse(strSections.Substring(strSections.LastIndexOf(' ') + 1)); // Get substring for total # of sections as an int


                LOG.WriteLine(DateTime.Now + "  : " + courseName + " - " + numSections + " available section(s)."); // Write to log 

                while (bMoreCourses) // While there are more courses
                {
                    try // Try to expand every class
                    {
                        if (isWheelGone()) // If the processing wheel is gone, click!
                        {
                            driver.FindElementById(courseID + section).Click(); // Expand every course through click command
                            LOG.WriteLine((courseID + section) + " click successful."); // Log click

                            section += 1; // Increment section
                        }

                        else // Problem if here
                        {
                            LOG.WriteLine(DateTime.Now);
                            LOG.WriteLine("ERROR: Problem occurred while trying to click classes");
                        }
                    }

                    catch (NoSuchElementException) // If here, no more classes were found
                    {
                        bMoreCourses = false; // Stop while loop
                    }
                }

                String pageSource = driver.PageSource; // Read entire course DOM of the page
                File.WriteAllText(courseName + "_DOM.txt", pageSource); // Write current course DOM to text file based off course

                LOG.WriteLine("\r\n\r\n" + DateTime.Now); // Write to log
                LOG.WriteLine(courseName + " expanded " + section + " section(s) successfully.\r\n\r\n\r\n"); // Write number of expanded sections

                driver.Navigate().Refresh(); // Refresh page and switch course
                driver.SwitchTo().Frame("ptifrmtgtframe"); // Switch to content frame
            }
        }

        [TestMethod]
        public void subjParser()
        {
            driver.SwitchTo().Frame("ptifrmtgtframe"); // Switch to content frame
            SelectElement subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list

            IList<IWebElement> subjectList = subjectDdList.Options; //Get list of IWeb elements
            int numSubjects = subjectList.Count; // Number of subjects CSUN has

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

            String subjName = "";

            for (int subjectIndex = 1; subjectIndex < numSubjects; subjectIndex++)
            {
                subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list
                subjectDdList.SelectByIndex(subjectIndex); // Select major

                if (isWheelGone())
                {
                    subjName = driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT").GetAttribute("value");

                    fixIllegalChar(ref subjName);

                    HtmlDocument doc = new HtmlDocument();
                    String filePath = @"F:\Projects\aspNet\DbPopulator\bin\Debug\" + subjName + "_DOM.txt";
                    doc.Load(filePath);

                    String coursesHtml = doc.GetElementbyId("PSCENTER").InnerText;
                    int numCourses = Int32.Parse(coursesHtml.Substring(coursesHtml.LastIndexOf(' ') + 1));

                    StreamWriter sw = new StreamWriter(subjName + "_parsed.txt");

                    for (int courseIndex = 0, courseRow = 0; courseIndex < numCourses; courseIndex++)
                    {
                        courseDescrHtml = getInnerText(doc, courseDescrID + courseIndex);
                        courseDescr = courseDescrHtml.Substring(0, courseDescrHtml.IndexOf('-')).Trim();

                        courseSecHtml = getInnerText(doc, courseSecID + courseIndex);
                        numSections = Int32.Parse(Regex.Match(courseSecHtml, @"\d+").Value);

                        //sw.WriteLine(courseDescrHtml + ": " + numSections + " section(s)");
                        for (int sectionIndex = 0; sectionIndex < numSections; sectionIndex++, courseRow++)
                        {
                            courseNumHtml = getInnerText(doc, courseNumID + courseRow);
                            courseNum = removeNewline(courseNumHtml);

                            courseLocHtml = getInnerText(doc, courseLocID + courseRow);
                            courseLoc = removeNewline(courseLocHtml);

                            courseDayHtml = getInnerText(doc, courseDayID + courseRow);
                            courseDay = removeNewline(courseDayHtml);
                            if (courseDay.Equals("&nbsp;"))
                                courseDay = "TBA";

                            courseTimeHtml = getInnerText(doc, courseTimeID + courseRow);
                            courseTime = removeNewline(courseTimeHtml);

                            courseInstrHtml = getInnerText(doc, courseInstrID + courseRow);
                            courseInstr = removeNewline(courseInstrHtml);

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
    }
}