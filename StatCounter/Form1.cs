using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace StatCounter
{

    public partial class frmMain : Form
    {
        //This is a boolean that stores whether or not there is a reference file. It will be assigned a value 
        //after the form is initialized.
        public bool referenceFileExists;

        public frmMain()
        {
            InitializeComponent();

            //TODO: What I'd like to do here is access a file that is used as a reference point.
            //This file would contain data collected from a previous file, so that more 
            //accurate statistics can be gathered. If no reference file exists, the first time a file
            //is loaded, that file will become the reference point. After a reference file is generated,
            //the next file loaded will become the reference file. Does that make sense??
            try
            {
                using (var fileStream = new FileStream("ref.txt", FileMode.Open))
                {
                    referenceFileExists = true;
                }
            }
            catch
            {
                MessageBox.Show("There appears to be no reference file. Load a file to create one.");
                referenceFileExists = false;
            }
        }

        OpenFileDialog ofd = new OpenFileDialog();

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                lblStats.Text = "";

                tbFileName.Text = ofd.FileName;
                Excel excel = new Excel(ofd.FileName, 1);


                //declaring these values and instantiating them all as 0
                double solids, cutstock, fj, core, backrip, waste;
                double percentTotal = 0;
                solids = cutstock = fj = core = backrip = waste = percentTotal;

                int backrip_quantity = 0;

                lblStats.Text = "LOADING...";

                //getting stats from the excel file
                getFileStats(ref solids, ref cutstock, ref fj, ref core, ref backrip, 
                    ref backrip_quantity, excel);

                /* --------------
                 * I no longer want the waste%. I want the waste bf volume
                 * 
                waste = Double.Parse(excel.ReadCell(46, 3)); // this is the waste%
                */

                waste = Double.Parse(excel.ReadCell(28, 3)); // this is the waste bf volume

                percentTotal = solids + cutstock + fj + core + backrip + waste;

                //displays the results onscreen
                printToScreen(waste, solids, cutstock, fj, core, backrip, backrip_quantity, percentTotal);

                excel.Close();
            }
        }

        //takes the stats and displays them in a label
        void printToScreen(double waste, double solids, double cutstock, double fj, 
            double core, double backrip, int backrip_quantity, double percentTotal)
        {
            lblStats.Text = "Waste: " + waste.ToString() + "%" + "\nSolids: " + solids.ToString() + "%" +
                    "\nCutstock: " + cutstock.ToString() + "%" +
                     "\nFJ: " + fj.ToString() + "%" + "\nCore: " + core.ToString() + "%" + "\nBackrip: " + backrip.ToString() + "%" +
                    "\nBackrip piece count: " + backrip_quantity.ToString() + " pcs" + "\n\nTOTAL: " + percentTotal.ToString() + "%";
        }

        //this gets the stats from an excel file through reference variables
        void getFileStats(ref double solids, ref double cutstock, ref double fj, ref double core, 
            ref double backrip, ref int backrip_quantity, Excel excel)
        {

            //ensuring that these are all set to zero
            solids = cutstock = fj = core = backrip = 0;

            int volPercentX = 0; //volume % x location
            int quantityX = 0;   //piece count x location
            int minLengthX = 0;  //min length x location
            int nameX = 0;       //name x location

            // this will run until they are all non-zero values. retrieves the x,y coordinates of the stats.
            // may have to improve this. the only reason i'm doing it this way is because the locations of
            // attributes can change. doesn't seem to take too long. 
            for (int i = 5; (volPercentX * quantityX * minLengthX * nameX) == 0; i++)
            {
                string name = excel.ReadCell(2, i);

                if (name == "Name")
                {
                    nameX = i;
                }
                else if (name == "Min. Length")
                {
                    minLengthX = i;
                }
                else if (name == "Quantity\n[ pcs ] ")
                {
                    quantityX = i;
                }

                else if (name == "Volume\n[ bf ] ")
                {
                    volPercentX = i;
                }

                //This section is commented out because it is using my previous (antiquated) method.
                //The method just above is the new method. 
                //I don't think I will use this one again, but I will leave it here as a reminder to myself
                /*
                else if (name == "Volume\n[ % ] ")
                {
                      volPercentX = i;
                }
                */

            }

            //this iterates through the actual values for each of the attributes until it reaches the end
            for (int i = 3; excel.ReadCell(i, minLengthX) != ""; i++)
            {
                double min_length = Double.Parse(excel.ReadCell(i, minLengthX));

                double volume = Double.Parse(excel.ReadCell(i, volPercentX));

                if (min_length < 8)
                {
                    string name = excel.ReadCell(i, nameX);
                    if (name.Contains("Core"))
                    {
                        core += volume;
                    }
                    else if (name.Contains("Backrip"))
                    {
                        backrip += volume;
                        backrip_quantity += int.Parse(excel.ReadCell(i, quantityX));
                    }
                    else
                    {
                        fj += volume;
                    }
                }
                else if (min_length <= 69 && min_length > 7)
                {
                    cutstock += volume;
                }
                else if (min_length <= 192 && min_length >= 69.5)
                {
                    solids += volume;
                }
            }
        }
    }
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int sheet)
        { 
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            var value = ws.Cells[i, j].Value2;
            if (value != null)
            {
                return value.ToString();
            }
            else
            {
                return "";
            }
        }

        public void Close()
        {
            this.excel.ActiveWorkbook.Close(0);
            this.excel.Quit();
        }
    }
}
