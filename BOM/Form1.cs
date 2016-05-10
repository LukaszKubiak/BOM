using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BOM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Open_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                CalculateSheet();
            }
        }

        private void CalculateSheet()
        {
            Excel.Application xlApp = new Excel.Application(); ;
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); ;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range range = xlWorkSheet.UsedRange;
            object misValue = System.Reflection.Missing.Value;

            string str;
            int rCnt = 0;
            int cCnt = 0;
            string[][] results = new string[range.Rows.Count - 16][];
            for (int i = 0; i < range.Rows.Count - 16; i++)
            {
                results[i] = new string[range.Rows.Count - 16];
            }

            List<string> lista = new List<string>();
            List<string> lista2 = new List<string>();

            range.UnMerge();
            for (rCnt = 18; rCnt < range.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= 4; cCnt++)
                {
                   str = (range.Cells[rCnt, cCnt] as Excel.Range).Text;
                    if (str == "" && (cCnt == 1 || cCnt == 2 || rCnt==18))
                        continue;
                    else
                        lista.Add(str);          
                }
            }
            rCnt = 18;
            cCnt = 1;
            str = (range.Cells[rCnt, cCnt] as Excel.Range).Text;
           
            while (rCnt < range.Rows.Count)
            {
                if (((range.Cells[rCnt, 1] as Excel.Range).Text == "" && (range.Cells[rCnt, 2] as Excel.Range).Text == ""))
                    rCnt++;

                lista2.Add((range.Cells[rCnt, 1] as Excel.Range).Text);
                lista2.Add((range.Cells[rCnt, 2] as Excel.Range).Text);
                if (!((range.Cells[rCnt, 3] as Excel.Range).Text == "" && (range.Cells[rCnt, 4] as Excel.Range).Text == ""))
                {
                    lista2.Add((range.Cells[rCnt, 3] as Excel.Range).Text);
                    lista2.Add((range.Cells[rCnt, 4] as Excel.Range).Text);
                }
                while ((range.Cells[++rCnt, 2] as Excel.Range).Text == "" && rCnt < range.Rows.Count)
                {
                    if ((range.Cells[rCnt, 3] as Excel.Range).Text == "" && (range.Cells[rCnt, 4] as Excel.Range).Text == "")
                        continue;
                    else {
                        lista2.Add((range.Cells[rCnt, 3] as Excel.Range).Text);
                        lista2.Add((range.Cells[rCnt, 4] as Excel.Range).Text);
                    }
                }

            }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}

