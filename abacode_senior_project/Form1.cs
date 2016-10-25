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
using System.Runtime.InteropServices;

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
            if (NessusRadio.Checked == true)
            {
                //----------------------------------
                // Ensure a path has been specified
                //----------------------------------
                if (pathTextBox.Text == "")
                {
                    MessageBox.Show("Please enter a path to a Nessus CSV file.");
                }
                else
                {
                    String path = Directory.GetCurrentDirectory();
                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = true;
                    excelApp.DisplayAlerts = false;
                    Excel.Workbook excelWorkbooks = null;
                    Excel.Workbook pivotTableTemplate = null;
                    

                    //-------------------------
                    // Check file extension
                    //-------------------------
                    if (Path.GetExtension(pathTextBox.Text).Equals(".csv"))
                    {
                        //-------------------------------------------------
                        // copy initial template for pivot table
                        //-------------------------------------------------
                        File.Copy("template.xlsx", pathTextBox.Text + "-parsedNessus.xlsx", true);
                        //-------------------------------------------------
                        // This opens the csv file and converts it to xlsx
                        // Saves it in the specified path in the text box
                        //-------------------------------------------------
                        try
                        {
                            /*  excelWorkbooks.OpenText(pathTextBox.Text,
                                      DataType: Excel.XlTextParsingType.xlDelimited,
                                      TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone,
                                      ConsecutiveDelimiter: true,
                                      Semicolon: true);*/
                            excelWorkbooks = excelApp.Workbooks.Open(pathTextBox.Text);

                            excelWorkbooks.SaveAs(pathTextBox.Text + ".xlsx", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                            excelWorkbooks.Close(); //closes first workbook

                            excelWorkbooks = excelApp.Workbooks.Open(pathTextBox.Text + ".xlsx");
                            Excel._Worksheet convertedCSVWorksheet = excelWorkbooks.ActiveSheet; //opens another workbook on index 1




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

                            /*
                             * transfer all the information
                             * into the excel pivot table
                             */
                             
                            
                            pivotTableTemplate = excelApp.Workbooks.Open(pathTextBox.Text + "-parsedNessus.xlsx");
                            Excel._Worksheet pivotTableData = pivotTableTemplate.Sheets[2];
                            

                            for (int i = 0; i < usedRange.Rows.Count; ++i)
                            {
                                if (i == 0)
                                {
                                    continue;
                                }
                                else
                                {
                                    //ID
                                    pivotTableData.Cells[i + 9, 1] = convertedCSVWorksheet.Cells[i + 1, 1];

                                    //IP
                                    pivotTableData.Cells[i + 9, 2] = convertedCSVWorksheet.Cells[i + 1, 5];

                                    //Hostname
                                    //No host name in nessus csv

                                    //Operating System
                                    //No operating system field in nessus csv

                                    //Vulnerability Name
                                    pivotTableData.Cells[i + 9, 5] = convertedCSVWorksheet.Cells[i + 1, 8];

                                    //Risk and Severity
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("None") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Info"))
                                    {
                                        pivotTableData.Cells[i + 9, 6] = "Info";
                                        pivotTableData.Cells[i + 9, 7] = "0";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Low"))
                                    {
                                        pivotTableData.Cells[i + 9, 6] = "Low";
                                        pivotTableData.Cells[i + 9, 7] = "1";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Medium"))
                                    {
                                        pivotTableData.Cells[i + 9, 6] = "Medium";
                                        pivotTableData.Cells[i + 9, 7] = "2";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("High"))
                                    {
                                        pivotTableData.Cells[i + 9, 6] = "High";
                                        pivotTableData.Cells[i + 9, 7] = "3";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Critical") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Severe"))
                                    {
                                        pivotTableData.Cells[i + 9, 6] = "Critical";
                                        pivotTableData.Cells[i + 9, 7] = "4";
                                    }

                                    //Service
                                    //No service field in nessus csv

                                    //Protocol
                                    pivotTableData.Cells[i + 9, 9] = convertedCSVWorksheet.Cells[i + 1, 6];

                                    //Port
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("0"))
                                    {
                                        pivotTableData.Cells[i + 9, 10] = "---";
                                    }
                                    else
                                    {
                                        pivotTableData.Cells[i + 9, 10] = convertedCSVWorksheet.Cells[i + 1, 7];
                                    }

                                    //Vulnerability description
                                    pivotTableData.Cells[i + 9, 11] = convertedCSVWorksheet.Cells[i + 1, 10];

                                    //Remediation
                                    pivotTableData.Cells[i + 9, 12] = convertedCSVWorksheet.Cells[i + 1, 11];

                                    //Results
                                    //No results field in nessus csv

                                    //Exploit available
                                    //Need to do this with web scraping

                                    //Vuln Publish Date
                                    //Web scrape this information as well

                                    //Patch Publish Date
                                    //Web scrap this

                                    //See Also
                                    pivotTableData.Cells[i + 9, 17] = convertedCSVWorksheet.Cells[i + 1, 12];

                                    //CVSS SCORE
                                    pivotTableData.Cells[i + 9, 18] = convertedCSVWorksheet.Cells[i + 1, 3];

                                    //CVSS vector
                                    //Web scrap this

                                    //CVE
                                    pivotTableData.Cells[i + 9, 20] = convertedCSVWorksheet.Cells[i + 1, 2];
                                }
                            }
                            //---------------------------------------------------
                            // Close worksheets and workbooks. Release Resources
                            // and close excel
                            //---------------------------------------------------
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(convertedCSVWorksheet);
                            excelWorkbooks.Close(false);
                            pivotTableTemplate.Close(true);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
                            excelApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                            MessageBox.Show("Done parsing file.");
                        }
                        catch(Exception err)
                        {
                            MessageBox.Show("Error encountered: could not open file. " + err);
                            if (excelWorkbooks != null)
                            {
                                excelWorkbooks.Close();
                            }
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
                            excelApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        }
                    }

                    else
                    {
                        if (excelWorkbooks != null)
                        {
                            excelWorkbooks.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbooks);
                        }
                        excelApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        MessageBox.Show("Please select a CSV file.");
                    }

                    
                    

                 }
            }

            //-----------------------
            // Nessus Report Parsing
            //-----------------------
            else
            {
                MessageBox.Show("OpenVAS Parsing is not implemented yet.");
            }
            this.Cursor = Cursors.Default;
        }
    }
}
