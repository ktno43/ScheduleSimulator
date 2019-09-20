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
using Oracle.DataAccess.Client;
using System.Data;
using System.Data.OleDb;

namespace DbPopulator
{
    [TestClass]
    public class DbPopulator
    {
        private static ChromeDriver driver = null; 
        private static StreamWriter swLog = null;
        private static StreamWriter swParseLog = null;
        private static String connString = "";

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
            connString = "User Id=kyle;Password=password;Data Source=kyle";
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

        private static String[] fixParse(String row) // Pass in a String row
        {
            String[] retString = new string[7];
            String[] rowSplit = row.Split(',');

            if (rowSplit.Length == 8) // If their is an instructor for the class
            {
                retString[0] = rowSplit[0];
                retString[1] = rowSplit[1];
                retString[2] = rowSplit[2];
                retString[3] = rowSplit[3];
                retString[4] = rowSplit[4];
                retString[5] = rowSplit[5];
                retString[6] = (rowSplit[6] + "," + rowSplit[7]).Replace("\"", string.Empty);
            }

            else // Else no instructor (STAFF)
            {
                retString[0] = rowSplit[0];
                retString[1] = rowSplit[1];
                retString[2] = rowSplit[2];
                retString[3] = rowSplit[3];
                retString[4] = rowSplit[4];
                retString[5] = rowSplit[5];
                retString[6] = rowSplit[6];
            }

            return retString; // Return string array
        }

        private static void createTable(OracleConnection connection, String tblName) // Create table in oracle
        {
            var commandCreateTbl = "CREATE TABLE " + tblName +
                " (COURSE VARCHAR2(30), " +
                "COURSENUM VARCHAR2(30), " +
                "LOCATION VARCHAR2(30), " +
                "DAYS VARCHAR2(30), " +
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
                    using (OracleCommand commandDelete = new OracleCommand(commandDelText, connection)) // Create command to delete table
                    {
                        commandDelete.Connection.Open();

                        try
                        {
                            commandDelete.ExecuteNonQuery(); // Execute command to delete table
                        }

                        catch (OracleException e) // No table was found 
                        {
                            String error = e.Message.ToString();

                            if (error.Contains("ORA-00942")) // No table found, so create a table
                                createTable(connection, tblName);
                        }

                        commandDelete.Connection.Close();
                    }

                    // Command to insert into oracle table
                    var commandText = String.Format("INSERT INTO {0} (Course,CourseNum,Location,Days,StartTime,EndTime,Instructor) " +
                        "VALUES(:Course,:CourseNum,:Location,:Days,:StartTime,:EndTime,:Instructor)", tblName);

                    using (OracleCommand command = new OracleCommand(commandText, connection)) // Command to insert into oracle database
                    {
                        // Add parameters
                        command.Parameters.Add(new OracleParameter(":Course", OracleDbType.Varchar2));
                        command.Parameters.Add(new OracleParameter(":CourseNum", OracleDbType.Varchar2));
                        command.Parameters.Add(new OracleParameter(":Location", OracleDbType.Varchar2));
                        command.Parameters.Add(new OracleParameter(":Days", OracleDbType.Varchar2));
                        command.Parameters.Add(new OracleParameter(":StartTime", OracleDbType.Varchar2));
                        command.Parameters.Add(new OracleParameter(":EndTime", OracleDbType.Varchar2));
                        command.Parameters.Add(new OracleParameter(":Instructor", OracleDbType.Varchar2));

                        // Add values to each parameters for each row in the file
                        while ((row = reader.ReadLine()) != null)
                        {
                            String[] arrRow = fixParse(row); // Parse the row

                            // Add values to the parameters
                            command.Parameters[":Course"].Value = arrRow[0];
                            command.Parameters[":CourseNum"].Value = arrRow[1];
                            command.Parameters[":Location"].Value = arrRow[2];
                            command.Parameters[":Days"].Value = arrRow[3];
                            command.Parameters[":StartTime"].Value = arrRow[4];
                            command.Parameters[":EndTime"].Value = arrRow[5];
                            command.Parameters[":Instructor"].Value = arrRow[6];

                            command.Connection.Open();
                            command.ExecuteNonQuery();
                            command.Connection.Close();
                        }
                    }
                }
            }
        }

        private String writeData(String subjWriteName, String subjTitle)
        {
            // Course
            String courseHtml = "";
            String courseClassID = "PSGRIDCOUNTER";
            int numCourses = 0;

            // Course section 
            String courseSecHtml = "";
            String courseSecID = "NR_SSS_SOC_NWRK_DESCR15$";
            int numSections = 0; // Number of sections

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

            String fileName = subjWriteName + "_PARSE.csv";
            String filePath = null;
            Boolean bError = false;

            try
            {
                using (var wCsv = new StreamWriter(fileName))
                {
                    courseHtml = driver.FindElementByClassName(courseClassID).Text; // String representation of number of sections offered in that course
                    numCourses = Int32.Parse(courseHtml.Substring(courseHtml.LastIndexOf(' ') + 1)); // Get number of courses under specified subject

                    swParseLog.Write(DateTime.Now); // parse log subject start
                    swParseLog.WriteLine("   " + subjTitle + ": " + numCourses + " course(s)"); // parse log subject name

                    for (int courseIndex = 0, courseRow = 0; courseIndex < numCourses; courseIndex++) // For the number of courses under the specified subject
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
                        }
                        swParseLog.WriteLine();
                    }
                    wCsv.Close();
                    swParseLog.WriteLine("\r\n\r\n");
                }
            }

            catch (Exception /* e */)
            {
                bError = true;
            }

            if (!bError)
            {
                filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
            }

            return filePath; // Return filepath
        }


        [TestMethod]
        public void scrapCourseDb()
        {
            driver.SwitchTo().Frame("ptifrmtgtframe"); // Switch to content frame
            SelectElement subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list

            IList<IWebElement> subjectList = subjectDdList.Options; //Get list of IWeb elements

            int numSubjects = subjectList.Count; // Number of subjects CSUN has
            String subjWriteName = ""; // Subject file name
            String subjTitle = ""; // Subject title
            String tblName = "tbl";

            String courseID = "SOC_DETAIL$"; // Prefix ID for courses

            Boolean bMoreCourses = true;
            int courseIndex = 0;

            for (int subjectIndex = 1; subjectIndex < numSubjects; subjectIndex++)
            {
                bMoreCourses = true;
                courseIndex = 0;

                subjectDdList = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")); // Select drop down list
                subjectDdList.SelectByIndex(subjectIndex); // Select subject

                subjWriteName = driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT").GetAttribute("value"); // Get the value of the selected item in the drop down list
                fixIllegalChar(ref subjWriteName); // Fix any illegal characters
                subjWriteName = Regex.Replace(subjWriteName, @"\s+", "");

                if (isWheelGone())
                {
                    driver.FindElementById("NR_SSS_SOC_NWRK_BASIC_SEARCH_PB").Click(); // Click the search button
                }

                else // Problem if here
                {
                    swLog.WriteLine(DateTime.Now);
                    swLog.WriteLine("ERROR IN " + subjTitle + ": Problem occurred while trying click search button.");
                }

                if (isWheelGone())
                {
                    subjTitle = new SelectElement(driver.FindElementById("NR_SSS_SOC_NWRK_SUBJECT")).SelectedOption.Text;

                    while (bMoreCourses)
                    {
                        try
                        {
                            if (isWheelGone())
                            {
                                driver.FindElementById(courseID + courseIndex).Click(); // Expand every course through click command
                                courseIndex += 1; // Increment course index
                            }

                            else // Problem if here
                            {
                                swLog.WriteLine(DateTime.Now);
                                swLog.WriteLine("ERROR IN " + subjTitle + ": Problem occurred while trying to expand courses.");
                            }
                        }

                        catch (NoSuchElementException) // If here, no more classes were found
                        {
                            bMoreCourses = false; // Stop while loop
                        }
                    }
                }

                else // Problem if here
                {
                    swLog.WriteLine(DateTime.Now);
                    swLog.WriteLine("ERROR IN " + subjTitle + ": Problem occurred while trying to load courses.");
                }

               String filePath = writeData(subjWriteName, subjTitle); // Pass subject file name and title of the subject and write the data

                if (filePath != null)
                {
                    insertOrcl(connString, filePath, (tblName + subjWriteName));
                    swLog.Write(DateTime.Now);
                    swLog.WriteLine("   " + "Finished creating table for subject: " + subjTitle + ".\r\n\r\n");
                }

                else
                {
                    swLog.Write(DateTime.Now);
                    swLog.WriteLine("   " + "ERROR attempting to creating table for subject: " + subjTitle + ".\r\n\r\n");
                }

                driver.Navigate().Refresh(); // Refresh page and switch course
                driver.SwitchTo().Frame("ptifrmtgtframe"); // Switch to content frame
            }
        }
    }
}
