using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace StatCounter
{

    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();

        }

        OpenFileDialog ofd = new OpenFileDialog();

        int fileCount = 0;

        Excel file1 = new Excel();
        Excel file2;

        public int volBFX;
        public int quantityX;
        public int minLengthX;
        public int nameX;

        private void GetXLocations(Excel file)
        {
            volBFX = 0;
            quantityX = 0;
            minLengthX = 0;
            nameX = 0;

            for (int i = 5; (volBFX * quantityX * minLengthX * nameX) == 0; i++)
            {
                string name = file.ReadCell(2, i);

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
                    volBFX = i;
                }
            }
        }
        private void BtnOpen_Click(object sender, EventArgs e)
        {

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                
                if (tbFile1.Text == "")
                {
                    lblStats.Text = "";
                    tbFile1.Text = ofd.FileName;
                }
                else
                {
                    tbFile2.Text = ofd.FileName;
                }

                if (file1.isVoid)
                {
                    file1 = new Excel(ofd.FileName, 1);
                }
                else
                {
                    file2 = new Excel(ofd.FileName, 1);
                }

                fileCount += 1;

                if (fileCount == 2)
                {
                    string run1 = file1.ReadCell(3, 2);
                    string run2 = file2.ReadCell(3, 2);

                    if (run1 == run2) //verifies that both files are part of the same run
                    {
                        double solids1 = 0;
                        double cutstock1 = 0;
                        double fj1 = 0;
                        double core1 = 0;
                        double backrip1 = 0;
                        double waste1 = 0;

                        double solids2 = 0;
                        double cutstock2 = 0;
                        double fj2 = 0;
                        double core2 = 0;
                        double backrip2 = 0;
                        double waste2 = 0;

                        double totalBF1 = 0;
                        double totalBF2 = 0;

                        int backrip_quantity1 = 0;
                        int backrip_quantity2 = 0;

                        lblStats.Text = "LOADING...";

                        GetXLocations(file1); //gets the locations of each of the columns from file 1

                        for (int i = 3; file1.ReadCell(i, 8) != ""; i++)
                        {
                            double min_length = Double.Parse(file1.ReadCell(i, minLengthX));

                            double volume;
                            volume = Double.Parse(file1.ReadCell(i, volBFX));

                            if (min_length < 8)
                            {
                                string name = file1.ReadCell(i, nameX);
                                if (name.Contains("Core"))
                                {
                                    core1 += volume;
                                }
                                else if (name.Contains("Backrip"))
                                {
                                    backrip1 += volume;
                                    backrip_quantity1 += int.Parse(file1.ReadCell(i, quantityX));
                                }
                                else
                                {
                                    fj1 += volume;
                                }
                            }
                            else if (min_length <= 69 && min_length > 7)
                            {
                                cutstock1 += volume;
                            }
                            else if (min_length <= 192 && min_length >= 69.5)
                            {
                                solids1 += volume;
                            }
                        }

                        waste1 = Double.Parse(file1.ReadCell(28, 3)); // this is the waste bf
                        totalBF1 = Double.Parse(file1.ReadCell(20, 3)); // total BF
                        

                        GetXLocations(file2); // gets the locations of each of the columns of file 2

                        for (int i = 3; file2.ReadCell(i, 8) != ""; i++)
                        {
                            double min_length = Double.Parse(file2.ReadCell(i, minLengthX));

                            double volume;
                            volume = Double.Parse(file2.ReadCell(i, volBFX));

                            if (min_length < 8)
                            {
                                string name = file2.ReadCell(i, nameX);
                                if (name.Contains("Core"))
                                {
                                    core2 += volume;
                                }
                                else if (name.Contains("Backrip"))
                                {
                                    backrip2 += volume;
                                    backrip_quantity2 += int.Parse(file2.ReadCell(i, quantityX));
                                }
                                else
                                {
                                    fj2 += volume;
                                }
                            }
                            else if (min_length <= 69 && min_length > 7)
                            {
                                cutstock2 += volume;
                            }
                            else if (min_length <= 192 && min_length >= 69.5)
                            {
                                solids2 += volume;
                            }
                        }

                        waste2 = Double.Parse(file2.ReadCell(28, 3)); // this is the waste BF
                        totalBF2 = Double.Parse(file2.ReadCell(20, 3)); // total BF

                        totalBF2 = totalBF2 - totalBF1;

                        double percentTotal = ((solids2 - solids1) + (cutstock2 - cutstock1) + (fj2 - fj1) + (core2 - core1) + (backrip2 - backrip1) + (waste2 - waste1)) / totalBF2;
                        // BE CAREFUL: REMEMBER -- this is where these are converted from BF to their percentages out of the total.

                        solids2 = (solids2 - solids1)/totalBF2;
                        cutstock2 = (cutstock2 - cutstock1)/totalBF2;
                        fj2 = (fj2 - fj1)/(totalBF2);
                        core2 = (core2 - core1)/totalBF2;
                        backrip2 = (backrip2 - backrip1)/totalBF2;
                        waste2 = (waste2 - waste1) / totalBF2;


                        backrip_quantity2 = backrip_quantity2 - backrip_quantity1;

                        lblStats.Text = "Waste: " + waste2.ToString("P") + "\nSolids: " + solids2.ToString("P") +
                            "\nCutstock: " + cutstock2.ToString("P") + "\nFJ: " + fj2.ToString("P") + "\nCore: " + core2.ToString("P")+ 
                            "\nBackrip: " + backrip2.ToString("P") + "\nBackrip piece count: " + backrip_quantity2.ToString() + " pcs" + 
                            "\n\nTOTAL: " + percentTotal.ToString("P");

                        file1.Close();
                        file2.Close();
                    }
                    else
                    {
                        MessageBox.Show("ERROR: Files are not from the same run.");
                    }
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

        public bool isVoid;

        public Excel(string path, int sheet)
        { 
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheet];

            isVoid = false;
        }

        public Excel()
        {
            this.path = null;
            wb = null;
            ws = null;

            isVoid = true;

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

        //public void SetVars(string path, int sheet)
        //{
        //    this.path = path;
        //    wb = excel.Workbooks.Open(path);
        //    ws = wb.Worksheets[sheet];
        //
         //   isVoid = false;
        //}
    }
}
