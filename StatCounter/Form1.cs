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

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                lblStats.Text = "";

                tbFileName.Text = ofd.FileName;
                Excel excel = new Excel(ofd.FileName, 1);

                double solids = 0;
                double cutstock = 0;
                double fj = 0;
                double core = 0;
                double backrip = 0;
                double waste = 0;
                double percentTotal = 0;

                int backrip_quantity = 0;

                lblStats.Text = "LOADING...";

                int volPercentX = 0; //volume % x location
                int quantityX = 0;   //piece count x location
                int minLengthX = 0;  //min length x location
                int nameX = 0;       //name x location

                // this will run until they are all non-zero values
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
                    else if (name == "Volume\n[ % ] ")
                    {
                        volPercentX = i;
                    }
                }
                for (int i = 3; excel.ReadCell(i, 8) != ""; i++)
                {
                    double min_length = Double.Parse(excel.ReadCell(i, minLengthX));

                    double volume;
                    volume = Double.Parse(excel.ReadCell(i, volPercentX));

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

                waste = Double.Parse(excel.ReadCell(46, 3)); // this is the waste%

                percentTotal = solids + cutstock + fj + core + backrip + waste;

                lblStats.Text = "Waste: " + waste.ToString() + "%" + "\nSolids: " + solids.ToString() + "%" +
                    "\nCutstock: " + cutstock.ToString() + "%" +
                     "\nFJ: " + fj.ToString() + "%" + "\nCore: " + core.ToString() + "%" + "\nBackrip: " + backrip.ToString() + "%" +
                    "\nBackrip piece count: " + backrip_quantity.ToString() + " pcs" + "\n\nTOTAL: " + percentTotal.ToString() + "%";

                excel.Close();
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
