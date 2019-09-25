﻿using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Oracle.DataAccess.Client;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics;

namespace CourseScraper
{
    class Program
    {
        private static ChromeDriver driver = null;
        private static StreamWriter swLog = null; // Database insert log
        private static StreamWriter swParseLog = null; // Parsed courses log

        const int SECOND = 1000;
        const int MILLISECOND = 1;

        // Loaded from my settings
        private static readonly String gChromeDriverDir = Properties.MySettings.Default.CHROME_DRIVER_DIRECTORY;
        private static readonly String gFileDir = Properties.MySettings.Default.FILE_DIRECTORY;
        private static readonly String gConnString = Properties.MySettings.Default.CONNECTION_STRING;


        static void Main(string[] args)
        {
            startCourseScrape();

            Console.WriteLine("Program Ending. . . ");
            Console.Write("Press ENTER to exit ----> ");
            Console.ReadLine();

            Environment.Exit(0);
        }

        private static void startCourseScrape()
        {
            initializeObj(); // Initialize settings

            Stopwatch stopWatch = new Stopwatch();

            stopWatch.Start();

            scrapCourseDb(); // Scrape course database

            stopWatch.Stop();

            TimeSpan ts = stopWatch.Elapsed;
            string elapsedTime = String.Format("{0:00}h {1:00}m {2:00}.{3:00}s",
                ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);

            logAndDisplay(swLog, "   Finished scraping in " + elapsedTime);

            cleanup(); // Cleanup drivers and processes 
        }

        private static void initializeObj()
        {
            driver = new ChromeDriver(gChromeDriverDir); // Initialize chrome driver
            driver.Navigate().GoToUrl(Properties.MySettings.Default.URL); // Navigate to URL
            driver.Manage().Window.Maximize(); // Maximize window
            swLog = new StreamWriter("log.txt"); // Log.txt
            swParseLog = new StreamWriter("log_parse.txt"); // log_parse.txt
        }

        private static void logAndDisplay(StreamWriter log, String msg)
        {
            // Write to console and log
            log.Write(DateTime.Now);
            log.WriteLine(msg);

            Console.Write(DateTime.Now);
            Console.WriteLine(msg);
        }

        private static void cleanup()
        {
            if (driver != null)
            {
                driver.Close(); // Close current browser window currently in focus

                Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver"); // Kill all chrome drivers

                foreach (var chromeDriverProcess in chromeDriverProcesses)
                {
                    chromeDriverProcess.Kill();
                }

                swLog.Close(); // Close the log
                swParseLog.Close(); // close the parse log

                driver = null;
            }
        }

        private static Boolean isWheelGone()
        {
            for (int i = 0; i < 300; i++) // Wait for a maximum of 30 seconds (300 * 100ms = 30,000 ms = 30 second)
            {
                Thread.Sleep(100 * MILLISECOND); // Wait for 1 ms

                IWebElement wheel = driver.FindElementById(Properties.MySettings.Default.PROCESSING_WHEEL); // Find processing wheel

                if (!isElementVisible(wheel)) // If wheel is NOT visible
                {
                    return true; // wheel is gone
                }
            }

            return false; // Problem at this point
        }

        private static Boolean isElementVisible(IWebElement element)
        {
            return element.Displayed && element.Enabled; // Return status of element being displayed and enabled
        }

        private static void fixIllegalChar(ref String refStr)
        {
            String invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars()); // Get a string of invalid characters

            foreach (char c in invalidChars)
            {
                refStr = refStr.Replace(c.ToString(), ""); // Remove invalid characters
            }
        }

        private static void fixCsvChar(ref String refData)
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

        private static String removeNewline(String str)
        {
            return str.Replace("\r\n", string.Empty); // Remove the newline in a string
        }

        private static String[] fixParse(String row) // Pass in a String row
        {
            String[] retString = new string[7]; // 7 colomns
            String[] rowSplit = row.Split(','); // Split at every comma (.CSV FILE) 

            if (rowSplit.Length == 8) // If their is an instructor for the class
            {
                retString[0] = rowSplit[0]; // Course section
                retString[1] = rowSplit[1]; // Course Number
                retString[2] = rowSplit[2]; // Course location
                retString[3] = rowSplit[3]; // Course days
                retString[4] = rowSplit[4]; // Course start time
                retString[5] = rowSplit[5]; // Course end time
                retString[6] = (rowSplit[6] + "," + rowSplit[7]).Replace("\"", string.Empty); // Course instructor
                // Remove double quotes and concatenate first & last name together
            }

            else // Else no instructor (STAFF)
            {
                retString[0] = rowSplit[0]; // Course section
                retString[1] = rowSplit[1]; // Course number
                retString[2] = rowSplit[2]; // Course location
                retString[3] = rowSplit[3]; // Course days
                retString[4] = rowSplit[4]; // course start time
                retString[5] = rowSplit[5]; // Course end time
                retString[6] = rowSplit[6]; // Course instructor
            }

            return retString; // Return string array
        }

        private static void writeSubj(IList<IWebElement> subjList)
        {
            using (var wSubj = new StreamWriter(Path.Combine(gFileDir, "SubjectList.txt")))
            {
                foreach (IWebElement subj in subjList) // For each element in the subject list, write down the subject name
                {
                    wSubj.WriteLine(subj.Text);
                }
            }
        }

        private static Boolean retryFindClick(By by)
        {
            Boolean result = false;
            int attempts = 0;

            while (attempts < 2)
            {
                try
                {
                    driver.FindElement(by).Click();
                    result = true;
                    break;
                }

                catch (StaleElementReferenceException /*e*/)
                {

                }
                attempts += 1;
            }

            return result;
        }

        private static void scrapCourseDb()
        {
            driver.SwitchTo().Frame(Properties.MySettings.Default.CONTENT_FRAME); // Switch to content frame
            SelectElement subjectDdList = new SelectElement(driver.FindElementById(Properties.MySettings.Default.SUBJECT_DDL)); // Select drop down list

            IList<IWebElement> subjectList = subjectDdList.Options; //Get list of IWeb elements
            writeSubj(subjectList); // Write all subjects to a file

            int numSubjects = subjectList.Count; // Number of subjects CSUN has
            String subjWriteName = ""; // Subject file name
            String subjTitle = ""; // Subject title

            String tblName = Properties.MySettings.Default.TBL_PREFIX; // tbl prefix

            String courseID = Properties.MySettings.Default.COURSE_ID; // Prefix ID for courses

            Boolean bMoreCourses = true;
            int courseIndex = 0;
            String clickElapsed = "";

            for (int subjectIndex = 1; subjectIndex < numSubjects; subjectIndex++)
            {
                Stopwatch swSubj = new Stopwatch();
                swSubj.Start();

                bMoreCourses = true;
                courseIndex = 0;

                subjectDdList = new SelectElement(driver.FindElementById(Properties.MySettings.Default.SUBJECT_DDL)); // Select drop down list
                subjectDdList.SelectByIndex(subjectIndex); // Select subject

                subjWriteName = driver.FindElementById(Properties.MySettings.Default.SUBJECT_DDL).GetAttribute("value"); // Get the value of the selected item in the drop down list
                fixIllegalChar(ref subjWriteName); // Fix any illegal characters
                subjWriteName = Regex.Replace(subjWriteName, @"\s+", "");

                if (isWheelGone())
                {
                    retryFindClick(By.Id(Properties.MySettings.Default.SEARCH_BTN)); // Click the search button

                    // driver.FindElementById(Properties.MySettings.Default.SEARCH_BTN).Click(); // Click the search button
                }

                else // Problem if here
                {
                    logAndDisplay(swLog, "   ERROR IN " + subjTitle + ": Problem occurred while trying click search button.");
                }

                if (isWheelGone())
                {
                    Stopwatch swClick = new Stopwatch();
                    subjTitle = new SelectElement(driver.FindElementById(Properties.MySettings.Default.SUBJECT_DDL)).SelectedOption.Text;

                    while (bMoreCourses)
                    {
                        try
                        {
                            swClick.Start();
                            if (isWheelGone())
                            {
                                retryFindClick(By.Id(courseID + courseIndex); // Expand every course through click command

                                // driver.FindElementById(courseID + courseIndex).Click(); // Expand every course through click command
                                courseIndex += 1; // Increment course index
                            }

                            else // Problem if here
                            {
                                logAndDisplay(swLog, "   ERROR IN " + subjTitle + ": Problem occurred while trying to expand courses.");
                            }
                        }

                        catch (NoSuchElementException) // If here, no more classes were found
                        {
                            bMoreCourses = false; // Stop while loop
                        }
                    }
                    swClick.Stop();
                    TimeSpan tsClick = swClick.Elapsed;
                    clickElapsed = String.Format("{0:00}m {1:00}.{2:00}s",
                        tsClick.Minutes, tsClick.Seconds, tsClick.Milliseconds / 10);
                }

                else // Problem if here
                {
                    logAndDisplay(swLog, "   ERROR IN " + subjTitle + ": Problem occurred while trying to load courses.");
                }


                Stopwatch swWrite = new Stopwatch();
                swWrite.Start();
                String filePath = writeData(subjWriteName, subjTitle); // Pass subject file name and title of the subject and write the data
                swWrite.Stop();
                TimeSpan tsWrite = swWrite.Elapsed;

                if (filePath != null)
                {
                    Stopwatch swInsert = new Stopwatch();
                    swInsert.Start();
                    insertOrcl(gConnString, filePath, (tblName + subjWriteName));
                    swInsert.Stop();
                    TimeSpan tsInsert = swInsert.Elapsed;
                    swSubj.Stop();

                    TimeSpan tsSubj = swSubj.Elapsed;
                    String subjElapsed = String.Format("{0:00}m {1:00}.{2:00}s",
                        tsSubj.Minutes, tsSubj.Seconds, tsSubj.Milliseconds / 10);

                    String insertElapsed = String.Format("{0:00}m {1:00}.{2:00}s",
                        tsInsert.Minutes, tsInsert.Seconds, tsInsert.Milliseconds / 10);

                    String writeElapsed = String.Format("{0:00}m {1:00}.{2:00}s",
                        tsWrite.Minutes, tsWrite.Seconds, tsWrite.Milliseconds / 10);


                    logAndDisplay(swLog, "   Finished inserting table for subject: " + subjTitle + " in " + subjElapsed);
                    swLog.Write(
                          "Expand Time: " + clickElapsed +
                          "\r\nWrite Time:  " + writeElapsed +
                          "\r\nInsert Time: " + insertElapsed +
                          "\r\n\r\n\r\n");
                }

                else
                {
                    logAndDisplay(swLog, "   ERROR attempting to creating table for subject: " + subjTitle + ".\r\n\r\n");
                }


                driver.Navigate().Refresh(); // Refresh page and switch course
                driver.SwitchTo().Frame(Properties.MySettings.Default.CONTENT_FRAME); // Switch to content frame
            }
        }

        private static String writeData(String subjWriteName, String subjTitle)
        {
            // Course
            String courseHtml = "";
            String courseClassID = Properties.MySettings.Default.COURSE_CLASS_ID;
            int numCourses = 0;

            // Course section 
            String courseSecHtml = "";
            String courseSecID = Properties.MySettings.Default.SECTION_ID;
            int numSections = 0; // Number of sections

            // Course number
            String courseNumHtml = "";
            String courseNum = "";
            String courseNumID = Properties.MySettings.Default.COURSE_NUM_ID;

            // Course location
            String courseLocHtml = "";
            String courseLoc = "";
            String courseLocID = Properties.MySettings.Default.COURSE_LOC_ID;

            // Course days
            String courseDayHtml = "";
            string courseDay = "";
            String courseDayID = Properties.MySettings.Default.COURSE_DAY_ID;

            // Course times
            String courseTimeHtml = "";
            String courseTime = "";
            String courseStartTime = "";
            String courseEndTime = "";
            String courseTimeID = Properties.MySettings.Default.COURSE_TIME_ID;

            // Course instructor
            String courseInstrHtml = "";
            String courseInstr = "";
            String courseInstrID = Properties.MySettings.Default.COURSE_INSTR_ID;

            // Course description
            String courseDescrHtml = "";
            String courseDescr = "";
            String courseDescrID = Properties.MySettings.Default.COURSE_DESCR_ID;
            String courseTitle = "";

            String fileName = subjWriteName + "_PARSE.csv";
            String filePath = null;
            Boolean bError = false;

            try
            {
                using (var wCsv = new StreamWriter(Path.Combine(gFileDir, fileName)))
                {
                    courseHtml = driver.FindElementByClassName(courseClassID).Text; // String representation of number of sections offered in that course
                    numCourses = Int32.Parse(courseHtml.Substring(courseHtml.LastIndexOf(' ') + 1)); // Get number of courses under specified subject

                    logAndDisplay(swParseLog, "   " + subjTitle + ": " + numCourses + " course(s)"); // parse log subject name

                    for (int courseIndex = 0, courseRow = 0; courseIndex < numCourses; courseIndex++) // For the number of courses under the specified subject
                    {
                        courseSecHtml = driver.FindElementById(courseSecID + courseIndex).Text; // Get the string of offered sections of that course
                        numSections = Int32.Parse(Regex.Match(courseSecHtml, @"\d+").Value); // Convert that string to a number

                        courseDescrHtml = driver.FindElementById(courseDescrID + courseIndex).Text;
                        courseTitle = courseDescrHtml.Substring(0, courseDescrHtml.IndexOf('(')).Trim();
                        courseDescr = courseDescrHtml.Substring(0, courseDescrHtml.IndexOf('-')).Trim(); // Trim the description to show only the class and section

                        swParseLog.WriteLine(courseTitle + ": " + numSections + " section(s)");
                        Console.WriteLine(courseTitle + ": " + numSections + " section(s)");

                        for (int sectionIndex = 0; sectionIndex < numSections; sectionIndex++, courseRow++) // For the number of sections within the class
                        {
                            courseNumHtml = driver.FindElementById(courseNumID + courseRow).Text; // Get the course number 
                            courseNum = removeNewline(courseNumHtml);

                            courseLocHtml = driver.FindElementById(courseLocID + courseRow).Text; // Get the location of the course
                            courseLoc = removeNewline(courseLocHtml);

                            courseDayHtml = driver.FindElementById(courseDayID + courseRow).Text; // Get the instruction days the class is taught
                            courseDay = removeNewline(courseDayHtml);
                            if (courseDay.Equals(" "))
                            {// If day is empty, the string is TBA
                                courseDay = "TBA";
                            }

                            courseTimeHtml = driver.FindElementById(courseTimeID + courseRow).Text; // Get the time the class is taught
                            courseTime = removeNewline(courseTimeHtml);

                            if (courseTime.Contains("-")) // If the class has a time
                            {
                                courseStartTime = courseTime.Substring(0, courseTime.IndexOf('-'));
                                courseEndTime = courseTime.Substring(courseTime.LastIndexOf('-') + 1);
                            }

                            else // Else no class time
                            {
                                courseStartTime = "TBA";
                                courseEndTime = "TBA";
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
                            Console.WriteLine("Section " + (sectionIndex + 1) + " of " + numSections + ": " + courseNum + " parsed successfully.");
                        }
                        swParseLog.WriteLine();
                        Console.WriteLine();
                    }
                    wCsv.Close();

                    swParseLog.WriteLine("\r\n\r\n");
                    Console.WriteLine("\r\n\r\n");
                }
            }

            catch (Exception /* e */)
            {
                bError = true;
            }

            if (!bError)
            {
                filePath = Path.Combine(gFileDir, fileName);
            }

            return filePath; // Return file path
        }

        private static void createOrclTbl(OracleConnection connection, String tblName) // Create table in oracle
        {
            var commandCreateTbl = "CREATE TABLE " + tblName +
                " (SECTION VARCHAR2(30), " +
                "COURSENUM VARCHAR2(30), " +
                "LOCATION VARCHAR2(30), " +
                "DAY VARCHAR2(30), " +
                "STARTTIME VARCHAR2(30), " +
                "ENDTIME VARCHAR2(30), " +
                "INSTRUCTOR VARCHAR2(100))";

            using (OracleCommand command = new OracleCommand(commandCreateTbl, connection))
            {
                command.ExecuteNonQuery();
            }
        }

        private static void insertOrcl(String connectionString, String filePath, String tblName) // Insert into oracle databse
        {
            using (var reader = new StreamReader(filePath)) // Open file
            {
                String row = "";
                var commandDelText = "DELETE FROM " + tblName;

                using (OracleConnection connection = new OracleConnection(connectionString)) // Establish oracle connection
                {
                    using (OracleCommand cmdDelete = new OracleCommand(commandDelText, connection)) // Create command to delete table
                    {
                        cmdDelete.Connection.Open();

                        try
                        {
                            cmdDelete.ExecuteNonQuery(); // Execute command to delete table
                        }

                        catch (OracleException e) // No table was found 
                        {
                            String error = e.Message.ToString();

                            if (error.Contains("ORA-00942")) // No table found, so create a table
                                createOrclTbl(connection, tblName);
                        }

                        cmdDelete.Connection.Close();
                    }

                    // Command to insert into oracle table
                    var commandText = String.Format("INSERT INTO {0} (Section,CourseNum,Location,Day,StartTime,EndTime,Instructor) " +
                        "VALUES(:Section,:CourseNum,:Location,:Day,:StartTime,:EndTime,:Instructor)", tblName);

                    using (OracleCommand cmdInsert = new OracleCommand(commandText, connection)) // Command to insert into oracle database
                    {
                        // Add parameters
                        cmdInsert.Parameters.Add(new OracleParameter(":Section", OracleDbType.Varchar2));
                        cmdInsert.Parameters.Add(new OracleParameter(":CourseNum", OracleDbType.Varchar2));
                        cmdInsert.Parameters.Add(new OracleParameter(":Location", OracleDbType.Varchar2));
                        cmdInsert.Parameters.Add(new OracleParameter(":Day", OracleDbType.Varchar2));
                        cmdInsert.Parameters.Add(new OracleParameter(":StartTime", OracleDbType.Varchar2));
                        cmdInsert.Parameters.Add(new OracleParameter(":EndTime", OracleDbType.Varchar2));
                        cmdInsert.Parameters.Add(new OracleParameter(":Instructor", OracleDbType.Varchar2));

                        // Add values to each parameters for each row in the file
                        while ((row = reader.ReadLine()) != null)
                        {
                            String[] arrRow = fixParse(row); // Parse the row

                            // Add values to the parameters
                            cmdInsert.Parameters[":Section"].Value = arrRow[0];
                            cmdInsert.Parameters[":CourseNum"].Value = arrRow[1];
                            cmdInsert.Parameters[":Location"].Value = arrRow[2];
                            cmdInsert.Parameters[":Day"].Value = arrRow[3];
                            cmdInsert.Parameters[":StartTime"].Value = arrRow[4];
                            cmdInsert.Parameters[":EndTime"].Value = arrRow[5];
                            cmdInsert.Parameters[":Instructor"].Value = arrRow[6];

                            cmdInsert.Connection.Open();
                            cmdInsert.ExecuteNonQuery();
                            cmdInsert.Connection.Close();
                        }
                    }
                }
            }
        }
    }
}
