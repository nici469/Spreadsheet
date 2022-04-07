using System;
using System.Collections.Generic;
using System.Text;

namespace Spreadsheet
{
    class ProcessString
    {
        /// <summary>
        /// converts the inputString into an array of its component lines
        /// </summary>
        /// <param name="inputString"></param>
        /// <returns></returns>
        string[] SeparateLines(string inputString)
        {
            //to store each separate line in the inputString as an element of an array
            string[] lines= new string[0];

            string fullString = inputString;
            
            //to store the number of newline characters in the input string
            int noOfNewLineChar = 0;
            
            //count the number of newline characters in the fullString
            for(int i = 0; i < fullString.Length; i++)
            {
                if (fullString[i] == '\n') noOfNewLineChar++;
            }

            
            //loop through all the lines in the fullString
            for(int linecount= noOfNewLineChar; linecount > 0; linecount--)
            {
                //in each iteration, search for the nearest newline character
                int nearestNwLnPos = 0;

                for(int i = 0; i < fullString.Length; i++)
                {
                    if (fullString[i] == '\n') { nearestNwLnPos = i;break; }
                }

                //extract the topmost line 
                string currrentLine=null;
                for(int i = 0; i < nearestNwLnPos; i++)
                {
                    currrentLine += fullString[i];
                }

                //cut out the any discovered line from fullString
                fullString = CutString(fullString, nearestNwLnPos);
                lines = AddLine(lines,currrentLine);

            }

            return lines;

        }

        /// <summary>
        /// cuts out portions of the data string from the first character up to and including the specified position
        /// and returns the new string
        /// </summary>
        /// <param name="data"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        string CutString(string stringData, int cutPos)
        {
            string outString = null;
            for(int i= cutPos + 1;i< stringData.Length; i++)
            {
                outString += stringData[i];
            }
            return outString;
        }

        /// <summary>
        /// returns a string array that is one element longer than the input string array, the new element being newLine
        /// </summary>
        /// <param name="lineArray"></param>
        /// <param name="newLine"></param>
        /// <returns></returns>
        string[] AddLine(string[] lineArray, string newLine)
        {
            //declare the output string array as one element longer than the input lineArray
            string[] outstringArray = new string[lineArray.Length + 1];

            //copy the elements of lineArray into the output string array
            for(int i = 0; i < lineArray.Length; i++)
            {
                outstringArray[i] = lineArray[i];
            }

            //add the new line to outstringArray
            outstringArray[lineArray.Length] = newLine;
            //return
            return outstringArray;
        }

        /// <summary>
        /// Converts an inputString into an array of strings using the specified "separator"
        /// </summary>
        /// <param name="inputString"></param>
        /// <param name="separator"></param>
        string[] SeparateLines(string inputString, char separator)
        {
            //to store each separate strings in the inputString as an element of an array
            string[] lines = new string[0];

            string fullString = inputString;

            //to store the number of "separator" characters in the input string
            int noOfSepChar = 0;

            //count the number of "separator" characters in the fullString or inputString
            for (int i = 0; i < fullString.Length; i++)
            {
                if (fullString[i] == separator) noOfSepChar++;
            }


            //loop through all the separate strings in the fullString
            for (int sepCount = noOfSepChar; sepCount > 0; sepCount--)
            {
                //in each iteration, search for the nearest "separator" character
                int nearestSepPos = 0;

                for (int i = 0; i < fullString.Length; i++)
                {
                    if (fullString[i] == separator) { nearestSepPos = i; break; }
                }

                //extract the topmost string 
                string currrentString = null;
                for (int i = 0; i < nearestSepPos; i++)
                {
                    currrentString += fullString[i];
                }

                //cut out the any discovered string from fullString
                fullString = CutString(fullString, nearestSepPos);
                lines = AddLine(lines, currrentString);

            }
            return lines;

        }


    }
}
