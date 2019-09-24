using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace tester
{
    class Program
    {
        static void Main(string[] args)
        {
            List <Course> cList = new List<Course>();
            String filePath = @"C:\Users\Jorg3\Desktop\School\CSUN\CompSci 490\UPDATED_COURSES_PARSED\COMP_PARSE.CSV";
            readFile(cList, filePath);
            //Course myCourse = new Course();
        }
        private static void readFile(List<Course> theCList, String filePath)
        {
            String row = "";
            
            using (var sr = new StreamReader(filePath))
            {
                while(sr.Peek() >= 0)
                {
                    row = sr.ReadLine();
                    String[] arrRow = fixParse(row);
                    Course myCourse = new Course(arrRow);
                    theCList.Add(myCourse);
                }
            }

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


    }
}
