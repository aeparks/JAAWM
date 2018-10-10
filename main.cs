/* Written by:      Aaron Parks
 * Assignment:      Capstone Project
 * Version:         2 (Beta)
 * Last Updated:    20 August 2012 @ 13:48
 * 
 * Project Update:
 * 
 * 
 * Project Description
 * Journeyman And Apprentice Workforce Manipulator (JAAWM) aka Capstone
 * =====================================================================
 * Purpose: This program serves two purposes.  One: validation of the information entered into the workforce
 * file.  The information in the workforce file needs to be consistant over the workforce file and the workforce
 * master record file in order for the proper diversity numbers to be generated. The Diversity Department at Sound
 * Transit maintains records of workforce demographics of the contractors they employ and the subcontractrs of
 * those contrators.  In order to qualify as a contractor or subcontractor for a Sound Transit project, a specific
 * workforce demographic is required.
 * Two: the program will act as an additional method to tabulate workforce information to ensure accuracy. Note
 * that this project does not replace the existing methods that the Diversity Department already uses.  Its
 * main purpose is error checking against the accepted methods. */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace jaawm_v2 {
   public partial class mainWindow : Form {
      //number of rows in current workforce file
      int w_rows = 0;
      //number of rows in current master file
      int m_rows = 0;
      //list of unique IDs
      List<String> idList = null;
      //for reading and manipulating the workforce fle
      private Excel.Application workforceFile = null;
      private Excel.Workbook workforceBook = null;
      private Excel.Sheets workforceSheets = null;
      //for reading and manipulating the master record file
      private Excel.Application masterFile = null;
      private Excel.Workbook masterBook = null;
      private Excel.Sheets masterSheets = null;

      public mainWindow() {
         InitializeComponent();
         this.killWorkforce();
         this.killMaster();
      }
      //---> 'workforceSelect_Click'
      /* Controls the open file dialog for selecting a workforce Excel file. Will perform error checking
       * after file is selected to ensure the workforce file is formatted appropriately. Once successfully
       * loaded, error checking on the record information (gender, A/J utilization, and EEO) and checking
       * on total sheet hours versus total weekly reported hours will be performed. If any error checks
       * fail, information about the failure will be displayed on the message panel. Also, any errors
       * encounterd must be fixed by the user. */
      private void workforceSelect_Click(object sender, EventArgs e) {
         //clear previous program state
         this.clearWindow();

         OpenFileDialog workforceDialog = new OpenFileDialog();
         //OpenFileDialog properties
         workforceDialog.InitialDirectory = "c:\\";
         workforceDialog.Filter = "Excel Files (*.xls; *.xlsx)|*.xls; *.xlsx";
         workforceDialog.FilterIndex = 2;
         workforceDialog.RestoreDirectory = true;

         //sentinel statement //will also perform error checking on file properties
         if (workforceDialog.ShowDialog() == DialogResult.OK) {
            this.msgOut.Text = "W O R K F O R C E    F I L E\n";
            //show file name in mainWindow
            this.workforceName.Text = Path.GetFileName(workforceDialog.FileName);
            workforceFile = new Excel.Application();

            //check if file is already open
            try {
               workforceBook = workforceFile.Workbooks.OpenXML(workforceDialog.FileName,
                  Type.Missing,
                  Type.Missing);
            }
            catch (Exception error_1) {
               this.msgOut.Text += "File Is Already Open!\nError:\n" + error_1 + "\n\n";
               this.killWorkforce();
               return;
            }
            this.msgOut.Text += "\"Open File\" Check Passed!\n";  //display check passed message
            workforceSheets = workforceBook.Worksheets;           //assign once checked for errors

            //check for proper tab formatting
            try {
               workforceSheets.get_Item(DateTime.Now.Year.ToString());
            }
            catch (Exception error_2) {
               this.msgOut.Text += "Improper Tab Name Formatting!\nError:\n" + error_2 + "\n\n";
               this.killWorkforce();
               return;
            }
            this.msgOut.Text += "\"Tab Name\" Check Passed!\n";   //display check passed message

            //verify gender, eeo, and a/j status in workforce file
            //will also calculate the number of records in the workforce file
            if (!this.verifyWorksheet(workforceSheets.get_Item(DateTime.Now.Year.ToString()))) {
               MessageBox.Show("Fix errors on workforce file " +
                  Path.GetFileName(workforceDialog.FileName) + " before continuing",
                  "Record Errors: " + Path.GetFileNameWithoutExtension(workforceDialog.FileName),
                  MessageBoxButtons.OK,
                  MessageBoxIcon.Information,
                  MessageBoxDefaultButton.Button1);
               this.killWorkforce();
            }
            else {
               this.msgOut.Text += "No Errors Found!\n\n";
               this.programLocation.Text = "Ready for Master Records File";
            }
            return;
         }
      }
      //---> 'masterSelect_Click'
      /* Controls the open file dialog for selecting the master record Excel file. Will perform error checking
       * after file is selected to ensure the master record file is formatted appropriately. Once successfully
       * loaded, error checking on the tab name will be performed. If the tab can't be found, an error message
       * stating the problem will be displayed to the message panel. */
      private void masterSelect_Click(object sender, EventArgs e) {
         //check if workforce file has been loaded first
         if (workforceFile == null) {
            MessageBox.Show("Must select and load Workforce file before selecting Master Record file.",
               "Master Record File Error",
               MessageBoxButtons.OK,
               MessageBoxIcon.Information,
               MessageBoxDefaultButton.Button1);
            return;
         }

         OpenFileDialog masterDialog = new OpenFileDialog();
         //OpenFileDialog properties
         masterDialog.InitialDirectory = "c:\\";
         masterDialog.Filter = "Excel Files (*.xls; *.xlsx)|*.xls; *.xlsx";
         masterDialog.FilterIndex = 2;
         masterDialog.RestoreDirectory = true;

         //sentinel statement //will also perform error checking on file properties
         if (masterDialog.ShowDialog() == DialogResult.OK) {
            //display on message panel
            this.msgOut.Text += "M A S T E R    F I L E\n";
            //show file name in mainWindow
            this.masterName.Text = Path.GetFileName(masterDialog.FileName);
            masterFile = new Excel.Application();

            //check if file is already open
            try {
               masterBook = masterFile.Workbooks.OpenXML(masterDialog.FileName,
                  Type.Missing,
                  Type.Missing);
            }
            catch (Exception error_3) {
               this.msgOut.Text += "File Is Already Open!\nError:\n" + error_3 + "\n\n";
               this.killMaster();
               return;
            }
            this.msgOut.Text += "\"Open File\" Check Passed!\n";  //display check passed message
            masterSheets = masterBook.Worksheets;                 //assign once checked for errors

            //check for proper tab formatting
            try {
               masterSheets.get_Item(this.masterTabName());
            }
            catch (Exception error_4) {
               this.msgOut.Text += "Tab Name \"" + this.masterTabName() + "\" Not Found!\nError:\n" +
                  error_4 + "\n\n";
               this.killMaster();
               return;
            }
            this.msgOut.Text += "\"Tab Name\" Check Passed!\n";      //display check passed message
         }
         this.msgOut.Text += "No Errors Found!\n\nReady to Merge Files\n";
         this.programLocation.Text = "Ready to Merge Files";
         this.progress.Text = String.Empty;
         return;
      }
      //---> 'mergeRecords_Click'
      /* Activates the main and final segment of this program. Will only run if valid workforce and master
       * record files have been loaded and passed all checks. There are two main parts to this function:
       * 1) a call to the 'copyRecords' function, passing in the current master and workforce Excel
       *    worksheets and the row number where records for the current year begin (description of function
       *    is at function declaration)
       * 2) a call to the 'parseNumbers' function, passing just the current master record file
       *    (description of function is at function declaration) */
      private void mergeRecords_Click(object sender, EventArgs e) {
         //pointers for shorter names
         Excel.Worksheet currentMaster = masterSheets.get_Item(this.masterTabName());
         Excel.Worksheet currentWorkforce = workforceSheets.get_Item(DateTime.Now.Year.ToString());

         //this.copyRecords(currentMaster, currentWorkforce, this.masterStartPoint(currentMaster));
         //this.parseNumbers(currentMaster);
         currentMaster.SaveAs("C:\\Users\\Jumbo\\Desktop\\Master " + DateTime.Now.Millisecond + ".xlsx",
            Type.Missing,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            Type.Missing);
         //update location and progress tracker
         this.programLocation.Text = "Done";
         this.progress.Text = String.Empty;
         //kill final processes
         this.killWorkforce();
         this.killMaster();
      }
      //---> 'dataVerify'
      /* Loops through columns H, I, and J checking for errors in the expected values of those columns. If 
       * an incorrect value is found, information about the error (location, data found, etc) will be 
       * displayed on the message panel. */
      private bool dataVerify(Excel.Worksheet current) {
         //update program location
         this.programLocation.Text = "Verifying Record Data";
         //create new stringbuilder for error messages
         StringBuilder errors = new StringBuilder("Data Errors\n");
         bool good = true;

         for (int r_temp = 3; r_temp < w_rows; r_temp++) {
            //check for valid record by checking employee ID#
            if (current.get_Range("F" + r_temp.ToString(), Type.Missing).Value != null) {
               //update progress tracker
               this.progress.Text = (Math.Round((double)r_temp / w_rows, 2) * 100).ToString() + "% Complete";
               //---> 1) Column H: gender -> M || F
               if ((current.get_Range("H" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "M") {
                  if ((current.get_Range("H" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "F") {
                     errors.Append("\nError: Row " + r_temp + "\nIncorrect Gender Char: \"" +
                        current.get_Range("H" + r_temp.ToString(), Type.Missing).Value + "\"\n");
                     good = false;
                  }
               }
               //---> 2) Column I: apprentice / journeyman progress -> 0% - 100%
               if (!(current.get_Range("I" + r_temp.ToString(), Type.Missing).Value <= 1 &&
                  current.get_Range("I" + r_temp.ToString(), Type.Missing).Value >= 0)) {
                  errors.Append("\nError: Row " + r_temp + "\nOut of Bounds Percentage: \"" +
                     (current.get_Range("I" + r_temp.ToString(), Type.Missing).Value * 100) + "%\"\n");
                  good = false;
               }
               //---> Column J: ethnic category -> ASI || BLK || CAU || HIS || NAT || OTH
               //check for missing EEO
               if (current.get_Range("J" + r_temp.ToString(), Type.Missing).Value == null) {
                  errors.Append("\nError: Row " + r_temp + "\nEEO Missing! (NULL Value)\n");
                  good = false;
               }
               if ((current.get_Range("J" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "CAU") {
                  if ((current.get_Range("J" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "ASI") {
                     if ((current.get_Range("J" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "BLK") {
                        if ((current.get_Range("J" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "HIS") {
                           if ((current.get_Range("J" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "NAT") {
                              if ((current.get_Range("J" + r_temp.ToString(), Type.Missing).Value).ToUpper() != "OTH") {
                                 errors.Append("\nError: Row " + r_temp + "\nIncorrect EEO Value: \"" +
                                    current.get_Range("J" + r_temp.ToString(), Type.Missing).Value + "\"\n");
                                 good = false;
                              }
                           }
                        }
                     }
                  }
               }
            }
         }
         //check which output to display
         if (good)
            this.msgOut.Text += "\"Data Verify\" Check Passed!\n";
         else
            this.msgOut.Text += errors.ToString();

         //clear program updates
         this.progress.Text = String.Empty;
         this.programLocation.Text = String.Empty;
         return good;
      }
      //---> 'verifyWorksheet'
      /* Checks the records of the parameter worksheet to verify correct options are put in the relavent categories
       * for gender, eeo, and a/j utilization. Will also compare the total calculated by Excel and the total of
       * weekly hours reported to check they are the same. If either check is false, then false will be returned.
       * If false is returned, the appropriate errors will be displayed on the message panel so that the user can
       * make the corrections to the workforce file. */
      private bool verifyWorksheet(Excel.Worksheet current) {
         w_rows = 3;                         //workforce records begin at the third row
         double sheetTotal = 0.0;            //the worksheet calculated total at the bottom of column k
         double calculatedTotal = 0.0;       //the total generated by 'tabulateHours' function
         bool totalComp = false;             //result of comparing 'sheetTotal' and 'calculatedHours'

         //---> 1) get number of records in the sheet
         this.programLocation.Text = "Tabulating Workforce Rows";
         for (int r_temp = 3; r_temp < current.UsedRange.Rows.Count; r_temp++) {
            //update progress tracker
            this.progress.Text = (Math.Round((double)r_temp / current.UsedRange.Rows.Count, 2) * 100).ToString() + "% Complete";
            if (current.get_Range("A" + r_temp.ToString(), Type.Missing).Value == null ||
               current.get_Range("A" + r_temp.ToString(), Type.Missing).Value == "Total")
               break;
            w_rows++;
         }
         //---> 2) get the calculated sheet total at the bottom of column k
         if (current.get_Range("K" + w_rows.ToString(), Type.Missing).Value != null)
            sheetTotal = current.get_Range("K" + w_rows.ToString(), Type.Missing).Value;
         else {
            //create a temp holder
            int w_temp = w_rows;
            while (true) {
               w_temp++;
               if (current.get_Range("K" + w_temp.ToString(), Type.Missing).Value != null) {
                  sheetTotal = current.get_Range("K" + w_temp.ToString(), Type.Missing).Value;
                  break;
               }
            }
         }
         sheetTotal = Math.Round(current.get_Range("K" + w_rows.ToString(), Type.Missing).Value, 2);
         //---> 3) get tabulated hours for the current sheet
         //calculatedTotal = this.tabulatedHours(current);

         //generate error message if totals don't match
         if (sheetTotal == calculatedTotal) {
            this.msgOut.Text += "\"Sheet Total\" Check Passed!\n";
            totalComp = true;
         }
         else
            this.msgOut.Text += "Sheet Total Error: Totals Do Not Match!\n";
         return totalComp && dataVerify(current);
      }
      //---> 'tabulateHours'
      /* Will sum all the numerical values for all the weeks reported so far in the workforce file. The
       * counting will be done across and down, meaning that all the hours reported so far for the
       * current year for each employee will be counted first before moving down to the next row. In
       * order to accelerate counting, only the weeks up to the current date will be counted. All 'null'
       * values after the current week will not be counted. Note: as the year progresses this method
       * will slow as more data needs to be tabulated. */
      private double tabulateHours(Excel.Worksheet current) {
         //update program location
         this.programLocation.Text = "Calculating Hours to Date";
         //number of weeks so far this year //will provide an end point for the function's read
         int countedColumns = (int)Math.Ceiling((double)DateTime.Now.DayOfYear / 7.0) + 13;
         //running total of summed hours
         double total = 0.0;

         for (int r_temp = 3; r_temp < w_rows; r_temp++) {
            //check for valid record by checking for employee ID#
            if (current.get_Range("F" + r_temp.ToString(), Type.Missing).Value != null) {
               for (int col = 13; col < countedColumns; col++) {
                  //update progress tracker
                  this.progress.Text = (Math.Round((double)r_temp / w_rows, 2) * 100).ToString() + "% Complete";
                  if (current.get_Range(convertHeader(col) + r_temp.ToString(), Type.Missing).Value is double)
                     total += current.get_Range(convertHeader(col) + r_temp.ToString(), Type.Missing).Value;
               }
            }
         }
         //clear progress tracker
         this.progress.Text = String.Empty;
         return Math.Round(total,2);
      }
      //---> 'copyRecords'
      /* Copies records from the workforce file to the appropriate worksheet on the Master Record file.
       * Not all records will be copied as doing so would be redundant in many cases. Instead, each
       * record in the workforce file (in order from top to bottom) will be compared to the existing
       * records in the Master Record file to check if it already exists. The various results from the
       * checks will determine what (if any) parts of the employee records are copied to the Master
       * Record file. */
      private void copyRecords(Excel.Worksheet master, Excel.Worksheet workforce, int m_start) {
         //need to check the state of 'separated' checkbox
         if (separated.Checked) {
            //update location tracker
            this.programLocation.Text = "Copying Record Data";
            //instantiate new string list for apprentice ids
            List<String> a_idList = new List<string>();
            //assign master pointer to first row of the current year
            int m_current = m_start;
            //assign workforce pointer to first record in workforce file
            int w_current = 2;

            //walk through master file and workforce file, copying updated hours information as necessary
            //if new record found in workforce, copy to master file
            for (; ; m_current++, w_current++) {
               //update progress tracker
               this.progress.Text = (Math.Round((double) w_current / w_rows, 2) * 100).ToString() + "% Complete";
               //check if records are the same; if so, update hours
               if (this.parityCheck(master, workforce, m_current, w_current)) {
                  master.get_Range("J" + m_current.ToString(), Type.Missing).Value =
                  workforce.get_Range("K" + w_current.ToString(), Type.Missing).Value;
               }
               else { //records are different in some way
               }
            } 
         }
         else { //checkbox is not checked
         }
      }

      //---> 'parityCheck
      /* Helper method to 'copyRecords'.  Performs a parity check against two records (defined by 'm_pointer' and
       * 'w_pointer' to determine if they are the same record. Records are compared by contractor, union, id number
       * craft, gender, a/j status, and eeo.  Only if all the checks are passed will true be returned. */
      private bool parityCheck(Excel.Worksheet master, Excel.Worksheet workforce, int m_pointer, int w_pointer) {
         return ((master.get_Range("B" + m_pointer.ToString(), Type.Missing).Value ==
            workforce.get_Range("A" + w_pointer.ToString(), Type.Missing).Value) &&
            (master.get_Range("C" + m_pointer.ToString(), Type.Missing).Value == 
            workforce.get_Range("B" + w_pointer.ToString(), Type.Missing).Value) &&
            (master.get_Range("E" + m_pointer.ToString(), Type.Missing).Value ==
            workforce.get_Range("F" + w_pointer.ToString(), Type.Missing).Value) &&
            (master.get_Range("F" + m_pointer.ToString(), Type.Missing).Value ==
            workforce.get_Range("G" + w_pointer.ToString(), Type.Missing).Value) &&
            (master.get_Range("G" + m_pointer.ToString(), Type.Missing).Value ==
            workforce.get_Range("H" + w_pointer.ToString(), Type.Missing).Value) &&
            (master.get_Range("H" + m_pointer.ToString(), Type.Missing).Value ==
            workforce.get_Range("I" + w_pointer.ToString(), Type.Missing).Value) &&
            (master.get_Range("I" + m_pointer.ToString(), Type.Missing).Value ==
            workforce.get_Range("J" + w_pointer.ToString(), Type.Missing).Value));
      }
      //---> 'parseNumbers'
      private void parseNumbers(Excel.Worksheet curent) {
      }
      //---> 'updateTotals'
      /* Pulls numbers from various cells to update the appropriate column total, row total, or
       * table total. Rather than call this function at the very end of the 'parseNumbers'
       * function this function is invoked after each pass sos that the totals update while the
       * cells update in the tables.
       * It was does this way purely for aesthetic reasons. */
      private void updateTotals() {
         //apprentice totals
         this.aCauTotal.Text = (Convert.ToInt32(this.amCau.Text) + Convert.ToInt32(this.afCau.Text)).ToString();
         this.aBlkTotal.Text = (Convert.ToInt32(this.amBlk.Text) + Convert.ToInt32(this.afBlk.Text)).ToString();
         this.aAsiTotal.Text = (Convert.ToInt32(this.amAsi.Text) + Convert.ToInt32(this.afAsi.Text)).ToString();
         this.aHisTotal.Text = (Convert.ToInt32(this.amHis.Text) + Convert.ToInt32(this.afHis.Text)).ToString();
         this.aNatTotal.Text = (Convert.ToInt32(this.amNat.Text) + Convert.ToInt32(this.afNat.Text)).ToString();
         this.aOthTotal.Text = (Convert.ToInt32(this.amOth.Text) + Convert.ToInt32(this.afOth.Text)).ToString();
         this.amTotal.Text = (Convert.ToInt32(this.amCau.Text) + Convert.ToInt32(this.amBlk.Text) +
             Convert.ToInt32(this.amAsi.Text) + Convert.ToInt32(this.amHis.Text) +
             Convert.ToInt32(this.amNat.Text) + Convert.ToInt32(this.amOth.Text)).ToString();
         this.afTotal.Text = (Convert.ToInt32(this.afCau.Text) + Convert.ToInt32(this.afBlk.Text) +
             Convert.ToInt32(this.afAsi.Text) + Convert.ToInt32(this.afHis.Text) +
             Convert.ToInt32(this.afNat.Text) + Convert.ToInt32(this.afOth.Text)).ToString();
         this.aTotal.Text = (Convert.ToInt32(this.amTotal.Text) + Convert.ToInt32(this.afTotal.Text)).ToString();
         //journeyman totals
         this.jCauTotal.Text = (Convert.ToInt32(this.jmCau.Text) + Convert.ToInt32(this.jfCau.Text)).ToString();
         this.jBlkTotal.Text = (Convert.ToInt32(this.jmBlk.Text) + Convert.ToInt32(this.jfBlk.Text)).ToString();
         this.jAsiTotal.Text = (Convert.ToInt32(this.jmAsi.Text) + Convert.ToInt32(this.jfAsi.Text)).ToString();
         this.jHisTotal.Text = (Convert.ToInt32(this.jmHis.Text) + Convert.ToInt32(this.jfHis.Text)).ToString();
         this.jNatTotal.Text = (Convert.ToInt32(this.jmNat.Text) + Convert.ToInt32(this.jfNat.Text)).ToString();
         this.jOthTotal.Text = (Convert.ToInt32(this.jmOth.Text) + Convert.ToInt32(this.jfOth.Text)).ToString();
         this.jmTotal.Text = (Convert.ToInt32(this.jmCau.Text) + Convert.ToInt32(this.jmBlk.Text) +
             Convert.ToInt32(this.jmAsi.Text) + Convert.ToInt32(this.jmHis.Text) +
             Convert.ToInt32(this.jmNat.Text) + Convert.ToInt32(this.jmOth.Text)).ToString();
         this.jfTotal.Text = (Convert.ToInt32(this.jfCau.Text) + Convert.ToInt32(this.jfBlk.Text) +
             Convert.ToInt32(this.jfAsi.Text) + Convert.ToInt32(this.jfHis.Text) +
             Convert.ToInt32(this.jfNat.Text) + Convert.ToInt32(this.jfOth.Text)).ToString();
         this.jTotal.Text = (Convert.ToInt32(this.jmTotal.Text) + Convert.ToInt32(this.jfTotal.Text)).ToString();
      }
      //---> 'masterStartPoint'
      /* Checks the cells in the 'A' column of the Master Record file ot find the first instnace
       * of either the current year or a blank cell. Finding the current year means find the
       * starting point to replace records. Finding a blank cell (null value) instead of finding
       * the current year means that no records from the current year have been inputted into
       * the Master Record file. */
      private int masterStartPoint(Excel.Worksheet current) {
         //update program location
         this.programLocation.Text = "Finding Starting Row";
         //master file begins on row 2
         int m_rows = 2;
         int r_counter = 0;
         for (; m_rows < current.UsedRange.Rows.Count; m_rows++) {
            //update progress tracker
            this.progress.Text = (Math.Round((double)m_rows / current.UsedRange.Rows.Count, 2) * 100).ToString() + "% Complete";
            //check for current year //insert records starting here
            if (current.get_Range("A" + m_rows.ToString(), Type.Missing).Value == DateTime.Now.Year &&
               current.get_Range("A" + (m_rows-1).ToString(), Type.Missing).Value != DateTime.Now.Year)
               r_counter = m_rows;
            if (current.get_Range("A" + m_rows.ToString(), Type.Missing).Value == null)
               break;
         }
         //clear progress field
         this.msgOut.Text += "Starting Rows: " + r_counter + "\n";
         this.msgOut.Text += "Total Rows: " + m_rows + "\n";
         this.progress.Text = String.Empty;
         return r_counter;
      }
      //---> 'masterTabName'
      /* Will try to extract the tab name on the 'Master Record File' from the name of the uploaded
       * workforce file. There are two know tab naming conventions: XXX-X.* and XXX-XX-X.*
       * This function will attempt to parse both naming convention types then hand back the string. */
      private string masterTabName() {
         string fileName = this.workforceName.Text;

         //split the string based upon deliminating characters '-' and '.'
         string[] fileNameArray = fileName.Split('-','.');

         //either [XXX][X][extension] or [XXX][XX][extension]
         if (fileNameArray[1].Length == 1)
            return fileNameArray[0];
         else
            return fileNameArray[0] + "-" + fileNameArray[1];
         return "\n";
      }
      //---> 'convertHeader'
      /* Converts an integer into an Excel column header (formatted as 'A', 'AB', 'ABC', etc.). It does
       * this by first converting the parameter into base 26 then converting the resulting integer into
       * the equivalent alphabetic letters by checking the highest order number placement, converting it
       * to a character, incrementing to the next number place, converting it to a character and then
       * appending it. */
      private string convertHeader(int columnNumber) {
         int dividend = columnNumber;
         string columnHeader = String.Empty;
         int modulo;

         while (dividend > 0) {
            modulo = (dividend - 1) % 26;
            columnHeader = Convert.ToChar(65 + modulo).ToString() + columnHeader;
            dividend = (int)((dividend - modulo) / 26);
         }
         return columnHeader;
      }
      //---> 'clearWindow'
      /* Once the program has completed a successful execution, the user may want to run it again before
       * closing the window. In that case, all labels and Excel objects must be cleared so that their
       * values don't interfer with the program's operation. This function will reset all the appropriate
       * labels and fields to their default value or a null value. */
      private void clearWindow() {
         //kill process (just in case)
         this.killWorkforce();
         this.killMaster();
         //reset file name labels
         this.workforceName.Text = this.masterName.Text = "File Name";
         //clear message panel ('msgpanel')
         this.msgOut.Text = String.Empty;
         //clear apprentice and journeyman labels
         this.amCau.Text = this.afCau.Text = this.aCauTotal.Text = this.amBlk.Text = this.afBlk.Text =
            this.aBlkTotal.Text = this.amAsi.Text = this.afAsi.Text = this.aAsiTotal.Text = this.amHis.Text =
            this.afHis.Text = this.aHisTotal.Text = this.amNat.Text = this.afNat.Text = this.aNatTotal.Text =
            this.amOth.Text = this.afOth.Text = this.aOthTotal.Text =
         this.jmCau.Text = this.jfCau.Text = this.jCauTotal.Text = this.jmBlk.Text = this.jfBlk.Text =
            this.jBlkTotal.Text = this.jmAsi.Text = this.jfAsi.Text = this.jAsiTotal.Text = this.jmHis.Text =
            this.jfHis.Text = this.jHisTotal.Text = this.jmNat.Text = this.jfNat.Text = this.jNatTotal.Text =
            this.jmOth.Text = this.jfOth.Text = this.jOthTotal.Text = this.amTotal.Text = this.afTotal.Text =
            this.aTotal.Text = this.jmTotal.Text = this.jfTotal.Text = this.jTotal.Text = "0";
      }
      //---> 'killWorkforce'
      /* Terminates all Excel processes associated with the Excel application instance 'workforce'. Garbage
       * collection must be called explicitly as the COM objects allocated by the Excel library are not
       * deleted automatically. Failure to force these processes to close will leave Excel processes running
       * on the system even after this program exits (memory leak). */
      private void killWorkforce() {
         if (workforceFile != null) {
            //first try/catch will cause 'workforceSheets' to be null thus causing an error
            if (workforceSheets != null) {
               Marshal.FinalReleaseComObject(workforceSheets);
               Marshal.FinalReleaseComObject(workforceBook);
            }
            workforceFile.Quit();
            Marshal.FinalReleaseComObject(workforceFile);
            //null dangling pointers
            workforceFile = null;
            workforceBook = null;
            workforceSheets = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
         }
      }
      //---> 'killMaster'
      /* Terminates all Excel processes associated with the Excel application instance 'master'. Garbage
       * collection must be called explicitly as the COM objects allocated by the Excel library are not
       * deleted automatically. Failure to force these processes to close will leave Excel processes running
       * on the system after this program exits (memory leak). */
      private void killMaster() {
         if (masterFile != null) {
            //first try/catch will cause 'masterSheets' to be null thus causing an error
            if (masterSheets != null) {
               Marshal.FinalReleaseComObject(masterSheets);
               Marshal.FinalReleaseComObject(masterBook);
            }
            masterFile.Quit();
            Marshal.FinalReleaseComObject(masterFile);
            //null dangling pointers
            masterFile = null;
            masterBook = null;
            masterSheets = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
         }
      }
   }
}