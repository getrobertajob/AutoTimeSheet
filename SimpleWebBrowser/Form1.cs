using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

/// <summary>
/// The author is Robert Lute
/// getrobertajob@gmail.com
/// 719-310-0055
/// 
/// The original intent of this application was for me to automatically log into my companies employee website, 
/// fill out the timesheet with values that are randomly adjusted values that are within tolerance level of company policy so that it doesn't always look the same, 
/// and then invoke the submit form attribute. 
/// 
/// I also setup a windows scheduled task to run this application on a weekly basis, 
/// so after writing this application I never had to actually fill out my timesheet and could focus on just work. 
/// 
/// Since I no longer work for that employer I can not log into that account I have pointed the application to a free timesheet site and do the same thing there to show how it is done. 
/// </summary>
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// This function Initializes the application
        /// </summary>
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// This funtion navigated the web browser object to the timesheet page
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://www.redcort.com/Free-Timecard-Calculator/");
        }

        /// <summary>
        /// This fuction is called once the webpage is finished loading
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            //Calls other fuctions to input the new data into the timesheet
            ChangeInHour();
            ChangeInMin();
            ChangeOutHour();
            ChangeOutMin();
            ChangeBreakTime();

            //Calls fuction to activate the print button on the page
            PrintTimeSheet();
        }

        /// <summary>
        /// This fuction is called when the about menu item is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void aboutMeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Displays a message box to show that I wrote the program
            MessageBox.Show("This program was made by Rober Lute");
        }

        /// <summary>
        /// This fuction is called when the exit menu item is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //This closes the application
            this.Close();
        }

        /// <summary>
        /// This function inputs the new data for begining hour fields column
        /// </summary>
        private void ChangeInHour()
        {
            //Declares variables, VarColNum=counter for loop, VarColName=string to hold name of column
            int VarColNum = 1;
            string VarColName = "start_hr";

            //Loop to cycle through only the first 5 in the html column
            while (VarColNum < 6)
            {
                //Appending the loop number to the column name
                VarColName += VarColNum;

                //Loop to cycle through html elements to find only the ones for the hour column and then modify the value
                //Used this method since the website doesn't have any elements that us an id
                foreach (HtmlElement he in webBrowser1.Document.All.GetElementsByName(VarColName))
                {
                    he.SetAttribute("value", "08");
                }

                //Adjusting variables, VarColNum=incremented to loop to next element, VarColName=setting back to base value so the next number can be appended at begining of loop
                VarColNum = ++VarColNum;
                VarColName = "start_hr";
            }
        }

        /// <summary>
        /// This function inputs the new data for begining minute fields column
        /// </summary>
        private void ChangeInMin()
        {
            //Declares variables, VarColNum=counter for loop, VarColName=string to hold name of column
            int VarColNum = 1;
            string VarColName = "start_min";
            Random rnd = new Random();

            //Loop to cycle through html elements to find only the ones for the hour column and then modify the value
            //Used this method since the website doesn't have any elements that us an id
            while (VarColNum < 6)
            {
                //Generate random value to adjust times
                int VarRndNum = rnd.Next(0, 6);
                string VarAdj = "0";

                //Appending random value to starting input string since you can't input an integer
                VarAdj += VarRndNum;

                //Appending the loop number to the column name
                VarColName += VarColNum;

                //Loop to cycle through html elements to find only the ones for the hour column and then modify the value
                //Used this method since the website doesn't have any elements that us an id
                foreach (HtmlElement he in webBrowser1.Document.All.GetElementsByName(VarColName))
                {
                    he.SetAttribute("value", VarAdj);
                }

                //Adjusting variables, VarColNum=incremented to loop to next element, VarColName=setting back to base value so the next number can be appended at begining of loop
                VarColNum = ++VarColNum;
                VarColName = "start_min";
            }
        }

        /// <summary>
        /// This function inputs the new data for ending hour fields column
        /// </summary>
        private void ChangeOutHour()
        {
            //Declares variables, VarColNum=counter for loop, VarColName=string to hold name of column
            int VarColNum = 1;
            string VarColName = "end_hr";

            //Loop to cycle through only the first 5 in the html column
            while (VarColNum < 6)
            {
                //Appending the loop number to the column name
                VarColName += VarColNum;

                //Loop to cycle through html elements to find only the ones for the hour column and then modify the value
                //Used this method since the website doesn't have any elements that us an id
                foreach (HtmlElement he in webBrowser1.Document.All.GetElementsByName(VarColName))
                {
                    he.SetAttribute("value", "05");
                }

                //Adjusting variables, VarColNum=incremented to loop to next element, VarColName=setting back to base value so the next number can be appended at begining of loop
                VarColNum = ++VarColNum;
                VarColName = "end_hr";
            }
        }

        /// <summary>
        /// This function inputs the new data for ending minute fields column
        /// </summary>
        private void ChangeOutMin()
        {
            //Declares variables, VarColNum=counter for loop, VarColName=string to hold name of column
            int VarColNum = 1;
            string VarColName = "end_min";
            Random rnd = new Random();

            //Loop to cycle through html elements to find only the ones for the hour column and then modify the value
            //Used this method since the website doesn't have any elements that us an id
            while (VarColNum < 6)
            {
                //Generate random value to adjust times
                int VarRndNum = rnd.Next(0, 6);
                string VarAdj = "0";

                //Appending random value to starting input string since you can't input an integer
                VarAdj += VarRndNum;

                //Appending the loop number to the column name
                VarColName += VarColNum;

                //Loop to cycle through html elements to find only the ones for the hour column and then modify the value
                //Used this method since the website doesn't have any elements that us an id
                foreach (HtmlElement he in webBrowser1.Document.All.GetElementsByName(VarColName))
                {
                    he.SetAttribute("value", VarAdj);
                }

                //Adjusting variables, VarColNum=incremented to loop to next element, VarColName=setting back to base value so the next number can be appended at begining of loop
                VarColNum = ++VarColNum;
                VarColName = "end_min";
            }
        }

        /// <summary>
        /// This function inputs the new data for break time fields column
        /// </summary>
        private void ChangeBreakTime()
        {
            //Declares variables, VarColNum=counter for loop, VarColName=string to hold name of column
            int VarColNum = 1;
            string VarColName = "break_min";

            //Loop to cycle through only the first 5 in the html column
            while (VarColNum < 6)
            {
                //Appending the loop number to the column name
                VarColName += VarColNum;

                //Loop to cycle through html elements to find only the ones for the hour column and then modify the value
                //Used this method since the website doesn't have any elements that us an id
                foreach (HtmlElement he in webBrowser1.Document.All.GetElementsByName(VarColName))
                {
                    he.SetAttribute("value", "30");
                }

                //Adjusting variables, VarColNum=incremented to loop to next element, VarColName=setting back to base value so the next number can be appended at begining of loop
                VarColNum = ++VarColNum;
                VarColName = "break_min";
            }
        }


        /// <summary>
        /// This function is called after the form has been updated and prints the results
        /// </summary>
        private void PrintTimeSheet()
        {
            //Gets a list of all the elements with the input tag
            HtmlElementCollection elc = this.webBrowser1.Document.GetElementsByTagName("input");

            //Loop to cycle through the list of elements with an input tag
            foreach (HtmlElement el in elc)
            {
                //Check to target the print button and invoke the click function
                if (el.GetAttribute("value").Equals("PRINT"))
                {
                    el.InvokeMember("Click");
                }
            }
        }
    }
}
