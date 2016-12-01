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
            if (NessusRadio.Checked == true)
            {
                browseDialog.InitialDirectory = "c:\\";
                browseDialog.Filter = "CSV (*.csv)|*.csv|XLSX (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                browseDialog.RestoreDirectory = true;
            }
            else
            {
                browseDialog.InitialDirectory = "c:\\";
                browseDialog.Filter = "XLSX (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|All files (*.*)|*.*";
                browseDialog.RestoreDirectory = true;

            }
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
                            excelWorkbooks = excelApp.Workbooks.Open(pathTextBox.Text);

                            excelWorkbooks.SaveAs(pathTextBox.Text + ".xlsx", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                            excelWorkbooks.Close();

                            excelWorkbooks = excelApp.Workbooks.Open(pathTextBox.Text + ".xlsx");
                            Excel._Worksheet convertedCSVWorksheet = excelWorkbooks.ActiveSheet;

                            //Get range to use for loops
                            Excel.Range usedRange = convertedCSVWorksheet.UsedRange;
                            
                            
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
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Critical") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 4].Value).Equals("Serious"))
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
                                    pivotTableData.Cells[i + 9, 8] = convertedCSVWorksheet.Cells[i + 1, 10];
                                    pivotTableData.Cells[i + 9, 9] = convertedCSVWorksheet.Cells[i + 1, 11];
                                    //cvss vector, exploit available, and vulnerability publish date
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value) != null && !(Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value).Equals("-")))
                                    {
                                        String[] cve = Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value).Split(','); //must do in openvas due to csv cve
                                        String NVDurl = "https://web.nvd.nist.gov/view/vuln/detail?vulnId=" + cve[0];
                                        SHDocVw.InternetExplorer IE = new SHDocVw.InternetExplorer();
                                        IE.Visible = false;
                                        IE.Navigate(NVDurl);
                                        System.Threading.Thread.Sleep(2000);
                                        mshtml.IHTMLDocument2 htmlDoc = IE.Document as mshtml.IHTMLDocument2;
                                        string content = htmlDoc.body.parentElement.outerHTML;
                                        int CVSSVectorStartIndex = content.IndexOf("(AV", content.IndexOf("(AV") + 1);
                                        int CVSSVectorEndIndex = 28;
                                        int vulnDateStartIndex = content.IndexOf("Original release date:") + 42;
                                        int vulnDateEndIndex = 10;
                                        pivotTableData.Cells[i + 9, 8] = convertedCSVWorksheet.Cells[i + 1, 10];
                                        pivotTableData.Cells[i + 9, 9] = convertedCSVWorksheet.Cells[i + 1, 11];
                                        if (vulnDateStartIndex != -1)
                                        {
                                            string stringVulndDate = content.Substring(vulnDateStartIndex, vulnDateEndIndex);
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
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 2].Value) != null)
                                    {
                                        pivotTableData.Cells[i + 9, 10] = "Exploit is available.";
                                    }

                                    //See Also
                                    pivotTableData.Cells[i + 9, 13] = convertedCSVWorksheet.Cells[i + 1, 12];

                                    //CVSS SCORE
                                    pivotTableData.Cells[i + 9, 14] = convertedCSVWorksheet.Cells[i + 1, 3];

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
                            MessageBox.Show("Done parsing file.");


                        }
                        catch (Exception err)
                        {
                            MessageBox.Show("Error encountered. Please try again.");
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


                            /*
                             * Getting range and iterating through
                             * the open file to get information
                             * for report
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
                                    if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 8].Value) != "n/a")
                                        pivotTableData.Cells[i + 4, 5] = convertedCSVWorksheet.Cells[i + 1, 8];
                                    else
                                        pivotTableData.Cells[i + 4, 5] = "Not available.";

                                    //Risk Level
                                    if (convertedCSVWorksheet.Cells[i + 1, 7].Value2 != null)
                                    {


                                        if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value2).Equals("None") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value2).Equals("Info"))
                                        {
                                            pivotTableData.Cells[i + 4, 6] = "Info";
                                            pivotTableData.Cells[i + 4, 7] = "0";
                                        }
                                        if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value2).Equals("Low"))
                                        {
                                            pivotTableData.Cells[i + 4, 6] = "Low";
                                            pivotTableData.Cells[i + 4, 7] = "1";
                                        }
                                        if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value2).Equals("Medium"))
                                        {
                                            pivotTableData.Cells[i + 4, 6] = "Medium";
                                            pivotTableData.Cells[i + 4, 7] = "2";
                                        }
                                        if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value2).Equals("High"))
                                        {
                                            pivotTableData.Cells[i + 4, 6] = "High";
                                            pivotTableData.Cells[i + 4, 7] = "3";
                                        }
                                        if (Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value2).Equals("Critical") || Convert.ToString(convertedCSVWorksheet.Cells[i + 1, 7].Value2).Equals("Serious"))
                                        {
                                            pivotTableData.Cells[i + 4, 6] = "Critical";
                                            pivotTableData.Cells[i + 4, 7] = "4";
                                        }
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
                                        System.Threading.Thread.Sleep(2000);
                                        mshtml.IHTMLDocument2 htmlDoc = IE.Document as mshtml.IHTMLDocument2;
                                        string content = htmlDoc.body.parentElement.outerHTML;
                                        int CVSSVectorStartIndex = content.IndexOf("(AV", content.IndexOf("(AV") + 1);
                                        int CVSSVectorEndIndex = 28;
                                        int vulnDateStartIndex = content.IndexOf("Original release date:") + 42;
                                        int vulnDateEndIndex = 10;
                                        if (vulnDateStartIndex != -1)
                                        {
                                            string stringVulndDate = content.Substring(vulnDateStartIndex, vulnDateEndIndex);
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
                            MessageBox.Show("Done parsing file.");
                        }
                        catch(Exception err)
                        {
                            MessageBox.Show("Error encountered. Please try again.");
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
                        MessageBox.Show("Please select a XLSX file.");
                    }
                }
            }
            this.Cursor = Cursors.Default;
        }
    }
}
