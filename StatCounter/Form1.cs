using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace StatCounter
{
    //TODO: I need to find a way to copy the stats loaded to the "Scanner Statistics" file.
    //This would save a little bit of effort on the part of the person loading the files

    public partial class frmMain : Form
    {
        //This is a boolean that stores whether or not there is a reference file. It will be assigned a value 
        //after the form is initialized.
        public bool referenceFileExists;

        public string referenceFile = "ref.txt"; // the name of the reference file

        public frmMain()
        {
            InitializeComponent();

            try
            {
                using (var fileStream = new FileStream(referenceFile, FileMode.Open))
                {
                    referenceFileExists = true;
                }
            }
            catch
            {
                MessageBox.Show("There appears to be no reference file. The first file loaded will create one.");
                referenceFileExists = false;
            }
        }

        OpenFileDialog ofd = new OpenFileDialog();

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                lblStats.Text = "";
                lblBackripPieces.Text = "";

                tbFileName.Text = ofd.FileName;
                Excel excel = new Excel(ofd.FileName, 1);


                //declaring these values and instantiating them all as 0
                double solids, cutstock, fj, core, backrip, waste;
                double volumeFromStats = 0;
                solids = cutstock = fj = core = backrip = waste = volumeFromStats;

                int backrip_quantity = 0;

                lblStats.Text = "LOADING...";

                //getting stats from the excel file
                getFileStats(ref solids, ref cutstock, ref fj, ref core, ref backrip, 
                    ref backrip_quantity, excel);


                waste = Double.Parse(excel.ReadCell(28, 3)); // this is the waste bf volume
                double total_volume = Double.Parse(excel.ReadCell(20, 3)); //this is the total volume from the run
                string run_date = excel.ReadCell(11, 2); //this is the run date of the file


                //this changes the way stats will be calculated based on whether or not a reference file exists
                if (referenceFileExists)
                {
                    using (StreamReader sr = new StreamReader(referenceFile))
                    {
                        String ref_text = sr.ReadToEnd();
                        string[] ref_vars = ref_text.Split(',');

                        //this indicates that the files being compared are from the same run
                        if (ref_vars[0] == run_date)
                        {

                            //getting all the stats from the reference file
                            double ref_solids = Double.Parse(ref_vars[1]);
                            double ref_cutstock = Double.Parse(ref_vars[2]);
                            double ref_fj = Double.Parse(ref_vars[3]);
                            double ref_core = Double.Parse(ref_vars[4]);
                            double ref_backrip = Double.Parse(ref_vars[5]);
                            int ref_backrip_quantity = int.Parse(ref_vars[6]); //don't think I'm going to use this one for anything yet
                            double ref_waste = Double.Parse(ref_vars[7]);
                            double ref_total_volume = Double.Parse(ref_vars[8]);


                            // What follows is a spectacularly clunky way to do this, but it should work

                            sr.Close();
                            writeReferenceFile(run_date, solids, cutstock, fj, core, backrip, backrip_quantity, waste, total_volume);


                            total_volume = (total_volume - ref_total_volume);

                            //updating the backrip quantity label here because it isn't always updated.
                            //will have to come up with a better way to do this
                            lblBackripPieces.Text = "Backrip from this shift: " + (backrip_quantity - ref_backrip_quantity).ToString();

                            solids = (solids - ref_solids);
                            cutstock = (cutstock - ref_cutstock);
                            fj = (fj - ref_fj);
                            core = (core - ref_core);
                            backrip = (backrip - ref_backrip);
                            waste = (waste - ref_waste);

                            volumeFromStats = solids + cutstock + fj + core + backrip + waste;

                            solids = solids / volumeFromStats;
                            cutstock = cutstock / volumeFromStats;
                            fj = fj / volumeFromStats;
                            core = core / volumeFromStats;
                            backrip = backrip / volumeFromStats;
                            waste = waste / volumeFromStats;

                            volumeFromStats = volumeFromStats / total_volume;


                        }
                        else
                        {
                            volumeFromStats = solids + cutstock + fj + core + backrip + waste;

                            MessageBox.Show("Files being compared are not from the same run. " +
                                "Legacy stats will be gathered and a new reference file will be created.");

                            sr.Close(); //I *think* that this closes the StreamReader so that I can write a new file. Seems to work.

                            writeReferenceFile(run_date, solids, cutstock, fj, core, backrip, backrip_quantity, waste, total_volume);

                            //I don't like using duplicate code, but this isn't *too* bad. I don't really want to create a method for this
                            //stats are gathered without a reference point
                            solids = solids / volumeFromStats;
                            cutstock = cutstock / volumeFromStats;
                            fj = fj / volumeFromStats;
                            core = core / volumeFromStats;
                            backrip = backrip / volumeFromStats;
                            waste = waste / volumeFromStats;
                            volumeFromStats = volumeFromStats / total_volume;

                        }

                    }
                }
                else
                {
                    volumeFromStats = solids + cutstock + fj + core + backrip + waste;

                    //write a new reference file
                    writeReferenceFile(run_date, solids, cutstock, fj, core, backrip, backrip_quantity, waste, total_volume);


                    //stats are gathered without a reference point
                    solids = solids / volumeFromStats;
                    cutstock = cutstock / volumeFromStats;
                    fj = fj / volumeFromStats;
                    core = core / volumeFromStats;
                    backrip = backrip / volumeFromStats;
                    waste = waste / volumeFromStats;
                    volumeFromStats = volumeFromStats / total_volume;
                }


                //displays the results onscreen
                printToScreen(waste, solids, cutstock, fj, core, backrip, backrip_quantity, volumeFromStats);

                excel.Close();
            }
        }

        //this writes a reference file for stat comparison
        void writeReferenceFile(string run_date, double solids, double cutstock, double fj, 
            double core, double backrip, int backrip_quantity, double waste, double total_volume)
        {
            using (StreamWriter writer = new StreamWriter(referenceFile, false))
            {
                writer.Write(run_date + "," + solids.ToString() + "," + cutstock.ToString() + "," + fj.ToString() + ","
                    + core.ToString() + "," + backrip.ToString() + "," + backrip_quantity.ToString() + ","
                    + waste.ToString() + "," + total_volume.ToString());

            }

            referenceFileExists = true;
        }

        //takes the stats and displays them in a label
        void printToScreen(double waste, double solids, double cutstock, double fj, 
            double core, double backrip, int backrip_quantity, double percentTotal)
        {
            lblStats.Text = "Waste: " + waste.ToString("p") + "\nSolids: " + solids.ToString("p") + "\nCutstock: " + cutstock.ToString("p") + 
                     "\nFJ: " + fj.ToString("p") + "\nCore: " + core.ToString("p") + "\nBackrip: " + backrip.ToString("p") +
                    "\nBackrip total piece count: " + backrip_quantity.ToString() + " pcs" + "\n\nTOTAL: " + percentTotal.ToString("p");
        }

        //this gets the stats from an excel file through reference variables
        void getFileStats(ref double solids, ref double cutstock, ref double fj, ref double core, 
            ref double backrip, ref int backrip_quantity, Excel excel)
        {

            //ensuring that these are all set to zero
            solids = cutstock = fj = core = backrip = 0;

            int vol_BF_X = 0; //volume % x location
            int quantityX = 0;   //piece count x location
            int minLengthX = 0;  //min length x location
            int nameX = 0;       //name x location

            // this will run until they are all non-zero values. retrieves the x,y coordinates of the stats.
            // may have to improve this. the only reason i'm doing it this way is because the locations of
            // attributes can change. doesn't seem to take too long. 
            for (int i = 5; (vol_BF_X * quantityX * minLengthX * nameX) == 0; i++)
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
                    vol_BF_X = i;
                }

            }

            //this iterates through the actual values for each of the attributes until it reaches the end
            for (int i = 3; excel.ReadCell(i, minLengthX) != ""; i++)
            {
                double min_length = Double.Parse(excel.ReadCell(i, minLengthX));

                double volume = Double.Parse(excel.ReadCell(i, vol_BF_X));

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
