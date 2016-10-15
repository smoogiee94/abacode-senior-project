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
    }
}
