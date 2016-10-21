using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace abacode_senior_project
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            //-----------------------------------
            // Create a Windows Open File Dialog
            // and sets the correct values
            //-----------------------------------
            OpenFileDialog browseDialog = new OpenFileDialog();
            browseDialog.InitialDirectory = "c:\\";
            browseDialog.Filter = "CSV (*.csv)|*.csv|XLSX (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            browseDialog.RestoreDirectory = true;

            //--------------------------------
            // If user hits ok on the dialog
            //--------------------------------
            if (browseDialog.ShowDialog() == DialogResult.OK)
            {
                pathTextBox.Text = browseDialog.FileName;
            }
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //-------------------------
            //openVAS Report Parsing
            //-------------------------
            if (openVASRadio.Checked == true)
            {
                //----------------------------------
                // Ensure a path has been specified
                //----------------------------------
                if (pathTextBox.Text == "")
                {
                    MessageBox.Show("Please enter a path to an openVAS file.");
                }
                else
                {
                    String path = Directory.GetCurrentDirectory();
                    var excelApp = new Excel.Application();
                    excelApp.DisplayAlerts = false;
                    var excelWorkbooks = excelApp.Workbooks;
                    

                    //-------------------------
                    // Check file extension
                    //-------------------------
                    if (Path.GetExtension(pathTextBox.Text).Equals(".csv"))
                    {
                        //-------------------------------------------------
                        // This opens the csv file and converts it to xlsx
                        // Saves it in the specified path in the text box
                        //-------------------------------------------------
                        try
                        {
                            excelWorkbooks.OpenText(pathTextBox.Text,
                                    DataType: Excel.XlTextParsingType.xlDelimited,
                                    TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone,
                                    ConsecutiveDelimiter: true,
                                    Semicolon: true);

                            excelWorkbooks[1].SaveAs(pathTextBox.Text + ".xlsx", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                            excelWorkbooks[1].Close(); //closes first workbook

                            
                            Excel._Worksheet convertedCSVWorksheet = excelWorkbooks.OpenXML(pathTextBox.Text + ".xlsx").ActiveSheet; //opens another workbook on index 1




                            //--------------------------------------
                            // File is now opened and ready. We must
                            // Now create a map and store all values
                            // for a vulnerability. Vulnerabilities
                            // are found by Plugin ID value
                            //--------------------------------------
                            /*
                             * Note: Use SortedDictionary with <int, List<String[]>> as the map
                             * Thus, we can use the Plugin ID and find all vulnerabilities
                             * of a given vulnerability
                             *
                             */

                            SortedDictionary<int, List<List<String>>> vulnerabilities = new SortedDictionary<int, List<List<String>>>();
                            List<String> vulnerabilityInformation;
                            Excel.Range usedRange = convertedCSVWorksheet.UsedRange;
                            for (int i = 0; i < usedRange.Rows.Count; ++i)
                            {
                                //---------------------------
                                // Skip first row of headers
                                //---------------------------
                                if (i == 0)
                                {
                                    continue;
                                }
                                else
                                {
                                    //------------------------------------
                                    // Create new list object to hold the
                                    // vulnerability information
                                    //------------------------------------
                                    vulnerabilityInformation = new List<String>();

                                    //--------------------------------------------
                                    // Iterate from cells 2-13 to gather 
                                    // vulnerability information from report
                                    //--------------------------------------------
                                    for (int j = 0; j < 13; ++j)
                                    {
                                        vulnerabilityInformation.Add(Convert.ToString(usedRange.Rows.Cells[i + 1, j + 1].Value));
                                    }

                                    //---------------------------------------------
                                    // Check to see if the current vulnerability
                                    // is in the map. If it is, add to the list
                                    // of that certain key
                                    //---------------------------------------------
                                    if (vulnerabilities.ContainsKey(Convert.ToInt32(usedRange.Rows.Cells[i + 1, 1].Value))){
                                        List<List<String>> vulnerabilityInformationCopy = new List<List<String>>();
                                        if(vulnerabilities.TryGetValue(Convert.ToInt32(usedRange.Rows.Cells[i + 1, 1].Value), out vulnerabilityInformationCopy))
                                        {
                                            //Must remove key
                                            vulnerabilities.Remove(Convert.ToInt32(usedRange.Rows.Cells[i + 1, 1].Value));

                                            //add information to the copy
                                            vulnerabilityInformationCopy.Add(vulnerabilityInformation);

                                            //add record back
                                            vulnerabilities.Add(Convert.ToInt32(usedRange.Rows.Cells[i + 1, 1].Value), vulnerabilityInformationCopy);
                                        }
                                    }
                                    //--------------------------------------
                                    // Else, create a new list<list<string>> 
                                    // object then add information then
                                    // add key
                                    //--------------------------------------
                                    else
                                    {
                                        List<List<String>> vulnerabilityInformationCopy = new List<List<String>>();
                                        vulnerabilityInformationCopy.Add(vulnerabilityInformation);
                                        vulnerabilities.Add(Convert.ToInt32(usedRange.Rows.Cells[i + 1, 1].Value), vulnerabilityInformationCopy);
                                    }
                                }
                            }

                            /*
                             * 
                             * Debugging statements. Was used to ensure maps were created correction
                             * TODO: Delete before final production
                            List<List<String>> foundVulnerability = new List<List<String>>();
                            if (vulnerabilities.TryGetValue(Convert.ToInt32(usedRange.Rows.Cells[3, 1].Value), out foundVulnerability))
                            {
                                for (int i = 0; i < foundVulnerability.Count; ++i)
                                {
                                    List<String> foundVulnerabilityInformation = foundVulnerability.ElementAt(i);
                                    for (int j = 0; j < foundVulnerabilityInformation.Count; ++j)
                                    {
                                        MessageBox.Show(foundVulnerabilityInformation.ElementAt(j));
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("No vulnerability of type" + Convert.ToInt32(usedRange.Rows.Cells[3, 1].Value) + "found");
                            }
                            */

                            //---------------------------------------------------
                            // Close worksheets and workbooks. Release Resources
                            // and close excel
                            //---------------------------------------------------
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(convertedCSVWorksheet);
                            excelWorkbooks[1].Close();
                            excelWorkbooks.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
                            excelApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        }
                        catch(Exception err)
                        {
                            MessageBox.Show("Error encountered: could not open file. " + err);
                            excelWorkbooks.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
                            excelApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        }
                    }

                    else
                    {
                        MessageBox.Show("Please select a CSV file.");
                    }

                    
                    

                 }
            }

            //-----------------------
            // Nessus Report Parsing
            //-----------------------
            else
            {

            }
            this.Cursor = Cursors.Default;
            MessageBox.Show("Done parsing file.");
        }
    }
}
