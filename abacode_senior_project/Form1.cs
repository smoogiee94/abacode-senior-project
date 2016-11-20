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
using System.Diagnostics;
using ScrapySharp.Network;
using ScrapySharp.Extensions;
using HtmlAgilityPack;
using System.Net;

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

        private void startButton_Click(object asdf, EventArgs e)
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
                        File.Copy("template_nessus.xlsx", pathTextBox.Text + "-parsedNessus.xlsx", true);
                        //-------------------------------------------------
                        // This opens the csv file and converts it to xlsx
                        // Saves it in the specified path in the text box
                        //-------------------------------------------------
                        try
                        {
                            //start timer for time analysis
                            Stopwatch stopwatch = new Stopwatch();
                            stopwatch.Start();

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

                            Excel.Range usedRange = convertedCSVWorksheet.UsedRange;
                            /*  SortedDictionary<int, List<List<String>>> vulnerabilities = new SortedDictionary<int, List<List<String>>>();
                              List<String> vulnerabilityInformation;
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
                              }*/

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

                                    //Vulnerability Name
                                    pivotTableData.Cells[i + 9, 3] = convertedCSVWorksheet.Cells[i + 1, 8];

                                    //Risk and Severity
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("None") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Info"))
                                    {
                                        pivotTableData.Cells[i + 9, 4] = "Info";
                                        pivotTableData.Cells[i + 9, 5] = "0";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Low"))
                                    {
                                        pivotTableData.Cells[i + 9, 4] = "Low";
                                        pivotTableData.Cells[i + 9, 5] = "1";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Medium"))
                                    {
                                        pivotTableData.Cells[i + 9, 4] = "Medium";
                                        pivotTableData.Cells[i + 9, 5] = "2";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("High"))
                                    {
                                        pivotTableData.Cells[i + 9, 4] = "High";
                                        pivotTableData.Cells[i + 9, 5] = "3";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Critical") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Severe"))
                                    {
                                        pivotTableData.Cells[i + 9, 4] = "Critical";
                                        pivotTableData.Cells[i + 9, 5] = "4";
                                    }

                                    //Protocol
                                    pivotTableData.Cells[i + 9, 6] = convertedCSVWorksheet.Cells[i + 1, 6];

                                    //Port
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("0"))
                                    {
                                        pivotTableData.Cells[i + 9, 7] = "---";
                                    }
                                    else
                                    {
                                        pivotTableData.Cells[i + 9, 7] = convertedCSVWorksheet.Cells[i + 1, 7];
                                    }

                                    //Vulnerability description
                                    pivotTableData.Cells[i + 9, 8] = convertedCSVWorksheet.Cells[i + 1, 10];

                                    //Remediation
                                    pivotTableData.Cells[i + 9, 9] = convertedCSVWorksheet.Cells[i + 1, 11];

                                    //Results
                                    //No results field in nessus csv

                                    //cvss vector 19
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value) != null && !(Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value).Equals("-")))
                                    {
                                        String[] cve = Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value).Split(','); //must do in openvas due to csv cve
                                        String NVDurl = "https://web.nvd.nist.gov/view/vuln/detail?vulnId=" + cve[0];
                                        SHDocVw.InternetExplorer IE = new SHDocVw.InternetExplorer();
                                        IE.Visible = false;
                                        IE.Navigate(NVDurl);
                                        System.Threading.Thread.Sleep(1000);
                                        mshtml.IHTMLDocument2 htmlDoc = IE.Document as mshtml.IHTMLDocument2;
                                        string content = htmlDoc.body.parentElement.outerHTML;
                                        int CVSSVectorStartIndex = content.IndexOf("(AV", content.IndexOf("(AV") + 1);
                                        int CVSSVectorEndIndex = 28;
                                        int vulnDateStartIndex = content.IndexOf("Original release date:") + 42;
                                        int vulnDateEndIndex = 10;
                                        if (vulnDateStartIndex != -1)
                                        {
                                            string stringVulndDate = content.Substring(vulnDateStartIndex, vulnDateEndIndex);
                                            Console.WriteLine(stringVulndDate);
                                            pivotTableData.Cells[i + 9, 11] = stringVulndDate;
                                        }
                                        if (CVSSVectorStartIndex != -1)
                                        {
                                            string stringCVSSVector = content.Substring(CVSSVectorStartIndex, CVSSVectorEndIndex);
                                            pivotTableData.Cells[i + 9, 15] = stringCVSSVector;
                                        }
                                        IE.Quit();
                                    }
                                    //Exploit available
                                    //Need to do this with web scraping
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value) != null)
                                    {
                                        pivotTableData.Cells[i + 9, 10] = "Exploit is available.";
                                    }
                                    //Vuln Publish Date
                                    //Web scrape this information as well

                                    //Patch Publish Date
                                    //Web scrap this

                                    //See Also
                                    pivotTableData.Cells[i + 9, 13] = convertedCSVWorksheet.Cells[i + 1, 12];

                                    //CVSS SCORE
                                    pivotTableData.Cells[i + 9, 14] = convertedCSVWorksheet.Cells[i + 1, 3];

                                    //CVSS vector
                                    //Web scrap this

                                    //CVE
                                    pivotTableData.Cells[i + 9, 16] = convertedCSVWorksheet.Cells[i + 1, 2];
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
                            //stop timer
                            stopwatch.Stop();
                            TimeSpan timespan = stopwatch.Elapsed;
                            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                                                                timespan.Hours,
                                                                timespan.Minutes,
                                                                timespan.Seconds,
                                                                timespan.Milliseconds / 10);

                            MessageBox.Show("Runtime: " + elapsedTime);
                            MessageBox.Show("Done parsing file.");


                        }
                        catch (Exception err)
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
                    if (Path.GetExtension(pathTextBox.Text).Equals(".xlsx"))
                    {
                        //-------------------------------------------------
                        // copy initial template for pivot table
                        //-------------------------------------------------
                        File.Copy("template_openvas.xlsx", pathTextBox.Text + "-parsedOpenVAS.xlsx", true);
                        try
                        {
                            //start timer for time analysis
                            Stopwatch stopwatch = new Stopwatch();
                            stopwatch.Start();
                            excelWorkbooks = excelApp.Workbooks.Open(pathTextBox.Text);
                            //                            excelWorkbooks.SaveAs(pathTextBox.Text + ".xlsx", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                            //                            excelWorkbooks.Close(); //closes first workbook

                            //                            excelWorkbooks = excelApp.Workbooks.Open(pathTextBox.Text + ".xlsx");
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

                            Excel.Range usedRange = convertedCSVWorksheet.UsedRange;
                            pivotTableTemplate = excelApp.Workbooks.Open(pathTextBox.Text + "-parsedOpenVAS.xlsx");
                            Excel._Worksheet pivotTableData = pivotTableTemplate.Sheets[2];


                            for (int i = 0; i < usedRange.Rows.Count; ++i)
                            {
                                if (i < 6)
                                {
                                    continue;
                                }
                                else
                                {
                                    //ID
                                    pivotTableData.Cells[i + 4, 1] = convertedCSVWorksheet.Cells[i + 1, 4];

                                    //IP
                                    pivotTableData.Cells[i + 4, 2] = convertedCSVWorksheet.Cells[i + 1, 2];

                                    //Hostname
                                    pivotTableData.Cells[i + 4, 3] = convertedCSVWorksheet.Cells[i + 1, 1];

                                    //Operating System/software check
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 13].Value) != "n/a")
                                    {
                                        pivotTableData.Cells[i + 4, 4] = convertedCSVWorksheet.Cells[i + 1, 13];
                                    }

                                    //Vulnerability Name
                                    pivotTableData.Cells[i + 4, 5] = convertedCSVWorksheet.Cells[i + 1, 8];

                                    //Risk Level
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("None") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("Info"))
                                    {
                                        pivotTableData.Cells[i + 4, 6] = "Info";
                                        pivotTableData.Cells[i + 4, 7] = "0";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("Low"))
                                    {
                                        pivotTableData.Cells[i + 4, 6] = "Low";
                                        pivotTableData.Cells[i + 4, 7] = "1";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("Medium"))
                                    {
                                        pivotTableData.Cells[i + 4, 6] = "Medium";
                                        pivotTableData.Cells[i + 4, 7] = "2";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("High"))
                                    {
                                        pivotTableData.Cells[i + 4, 6] = "High";
                                        pivotTableData.Cells[i + 4, 7] = "3";
                                    }
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("Critical") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value).Equals("Severe"))
                                    {
                                        pivotTableData.Cells[i + 4, 6] = "Critical";
                                        pivotTableData.Cells[i + 4, 7] = "4";
                                    }

                                    String[] parenthSplit = Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 3].Value).Split('(');
                                    //service
                                    pivotTableData.Cells[i + 4, 8] = parenthSplit[0];
                                    //protocol
                                    String[] slashSplit = parenthSplit[1].Split('/');
                                    pivotTableData.Cells[i + 4, 9] = slashSplit[1].Substring(0, slashSplit[1].Length - 1);
                                    //port
                                    pivotTableData.Cells[i + 4, 10] = slashSplit[0];

                                    //Vulnerability Description
                                    pivotTableData.Cells[i + 4, 11] = convertedCSVWorksheet.Cells[i + 1, 9];

                                    //Remediation check
                                    if (!(Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 10].Value2).Equals("n/a")))
                                    {
                                        pivotTableData.Cells[i + 4, 12] = convertedCSVWorksheet.Cells[i + 1, 10];
                                    }

                                    //Results check
                                    // pivotTableData.Cells[i + 4, 13] = convertedCSVWorksheet.Cells[i + 1, 12];
                                    if (!(Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 12].Value2).Equals("n/a")))
                                    {
                                        pivotTableData.Cells[i + 4, 13] = convertedCSVWorksheet.Cells[i + 1, 12].Value2;
                                    }
                                    //pivotTableData.Cells[i + 4, 13] = convertedCSVWorksheet.Cells[i + 1, 12].Value2;

                                    //exploit available
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 6].Value) != null && !(Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 6].Value).Equals("-")))
                                    {
                                        pivotTableData.Cells[i + 4, 14] = "Exploit is available.";
                                        //cvss score 18
                                        if (Convert.ToInt32(convertedCSVWorksheet.Cells[i + 1, 5].Value) != 0)
                                        {
                                            pivotTableData.Cells[i + 4, 18] = convertedCSVWorksheet.Cells[i + 1, 5];
                                        }

                                        //cvss vector 19
                                        String[] cve = Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 6].Value).Split(','); //must do in openvas due to csv cve
                                        String NVDurl = "https://web.nvd.nist.gov/view/vuln/detail?vulnId=" + cve[0];
                                        pivotTableData.Cells[i + 4, 17] = NVDurl;
                                        SHDocVw.InternetExplorer IE = new SHDocVw.InternetExplorer();
                                        IE.Visible = false;
                                        IE.Navigate(NVDurl);
                                        System.Threading.Thread.Sleep(1000);
                                        mshtml.IHTMLDocument2 htmlDoc = IE.Document as mshtml.IHTMLDocument2;
                                        string content = htmlDoc.body.parentElement.outerHTML;
                                        int CVSSVectorStartIndex = content.IndexOf("(AV", content.IndexOf("(AV") + 1);
                                        int CVSSVectorEndIndex = 28;
                                        int vulnDateStartIndex = content.IndexOf("Original release date:") + 42;
                                        int vulnDateEndIndex = 10;
                                        if (vulnDateStartIndex != -1)
                                        {
                                            string stringVulndDate = content.Substring(vulnDateStartIndex, vulnDateEndIndex);
                                            Console.WriteLine(stringVulndDate);
                                            pivotTableData.Cells[i + 4, 15] = stringVulndDate;
                                        }
                                        if (CVSSVectorStartIndex != -1)
                                        {
                                            string stringCVSSVector = content.Substring(CVSSVectorStartIndex, CVSSVectorEndIndex);
                                            pivotTableData.Cells[i + 4, 19] = stringCVSSVector;
                                        }
                                        IE.Quit();

                                        //cve 20
                                        pivotTableData.Cells[i + 4, 20] = convertedCSVWorksheet.Cells[i + 1, 6];

                                        //Console.WriteLine(cve[0]);
                                        //WebClient browser = new WebClient();
                                        //string nvdWebsite = browser.DownloadString(NVDurl);
                                        //int index1 = nvdWebsite.IndexOf("Date Entry Created") + 49;
                                        //int index2 = 8;
                                        //string substr = nvdWebsite.Substring(index1, index2);
                                        //Console.WriteLine(substr);
                                        //HtmlAgilityPack.HtmlDocument page = new HtmlAgilityPack.HtmlDocument();
                                        //Console.WriteLine(nvdWebsite);
                                        //WatiN.Core.Settings.MakeNewIeInstanceVisible = false;
                                        //WatiN.Core.Settings.AutoMoveMousePointerToTopLeft = false;
                                        //page.LoadHtml(nvdWebsite);
                                        //HtmlNode titleNode = page.DocumentNode.SelectSingleNode("/html/body/div[1]/div[4]/table/tr[9]/td/b");
                                        //if (titleNode == null)
                                        //{
                                        //    Console.WriteLine("NULL OMG NULL");
                                        //}
                                        //HtmlNode titleNode = pageResult.Html.CssSelect(".hohoho-header").First();
                                        //string stuffnode = titleNode.InnerText;
                                        //Console.WriteLine(titleNode);
                                        //pivotTableData.Cells[i + 4, 15] = titleNode;

                                    }

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
                }
            }
            this.Cursor = Cursors.Default;
        }
    }
}
