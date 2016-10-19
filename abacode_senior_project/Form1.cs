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
            Stream inputStream = null;
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
//                                    MessageBox.Show(usedRange.Rows.Cells[i + 1, 1].Value.ToString());
                                }
                            }



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
                        catch (Exception err)
                        {
                            MessageBox.Show("Error encountered: " + err);
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
        }
    }
}
