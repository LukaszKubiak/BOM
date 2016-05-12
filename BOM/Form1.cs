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

            int rCnt = 0;
            string[][] results = new string[range.Rows.Count - 16][];
            List<string> lista = new List<string>();
            List<string> lista2 = new List<string>();
            List<Item> items = new List<Item>();
            List<Item> children = new List<Item>();
            for (rCnt = 18; rCnt <= range.Rows.Count;)
            {
                var item = new Item();
                item.Line = "H";
                item.ItemCode = (range.Cells[rCnt, 1] as Excel.Range).Text;
                item.ItemDesc = (range.Cells[rCnt, 2] as Excel.Range).Text;
                string childCode = (range.Cells[rCnt, 3] as Excel.Range).Text;
                string childDesc = (range.Cells[rCnt, 4] as Excel.Range).Text;
                string childQuantity = (range.Cells[rCnt, 7] as Excel.Range).Text;
                if (childCode != "" && childDesc != "")
                    item.Children.Add(new Child(childCode, childDesc, childQuantity));
                int quantityIndex = rCnt;
            while ((range.Cells[++rCnt, 2] as Excel.Range).Text == "" && rCnt <= range.Rows.Count)
                {
                    childCode = (range.Cells[rCnt, 3] as Excel.Range).Text;
                    childDesc = (range.Cells[rCnt, 4] as Excel.Range).Text;
                    childQuantity = (range.Cells[rCnt, 7] as Excel.Range).Text;
                    if(childCode == "" && childDesc == "" && !(range.Cells[rCnt, 3] as Excel.Range).MergeCells)
                    {
                        childCode = (range.Cells[rCnt, 11] as Excel.Range).Text;
                        childDesc = (range.Cells[rCnt, 12] as Excel.Range).Text;
                        childQuantity = (range.Cells[rCnt, 14] as Excel.Range).Text;
                    }
                    if (childQuantity == "")
                        childQuantity = (range.Cells[quantityIndex, 7] as Excel.Range).Text;
                    else
                        quantityIndex = rCnt;
                    if(childCode != "" && childDesc != "")
                        item.Children.Add(new Child(childCode, childDesc, childQuantity));
                }
                items.Add(item);

            }
            for (int i = 0; i < items.Count; i++)
            {
                foreach (var child in items[i].Children)
                {
                    var item = new Item();
                    item.Line = "H";
                    item.ItemCode = child.ItemCode;
                    item.ItemDesc = child.ItemDesc;
                    for (rCnt = 18; rCnt <= range.Rows.Count;rCnt++)
                    {
                        if ((range.Cells[rCnt, 3] as Excel.Range).Text == item.ItemCode)
                        {
                            string childCode = (range.Cells[rCnt, 11] as Excel.Range).Text;
                            string childDesc = (range.Cells[rCnt, 12] as Excel.Range).Text;
                            string childQuantity = (range.Cells[rCnt, 14] as Excel.Range).Text;
                            if (childCode != "" && childDesc != "" && childCode != item.ItemCode)
                                item.Children.Add(new Child(childCode, childDesc, childQuantity));
                            while ((range.Cells[++rCnt, 3] as Excel.Range).Text == "" && rCnt <= range.Rows.Count && (range.Cells[rCnt, 3] as Excel.Range).MergeCells)
                            {
                                childCode = (range.Cells[rCnt, 11] as Excel.Range).Text;
                                childDesc = (range.Cells[rCnt, 12] as Excel.Range).Text;
                                childQuantity = (range.Cells[rCnt, 14] as Excel.Range).Text;
                                if(childCode != "" && childDesc != "")
                                    item.Children.Add(new Child(childCode, childDesc, childQuantity));
                            }
                            break;
                        }
                    }
                    children.Add(item);
                }
                
            }
            foreach (var item in items)
            {
                
                    textBox1.AppendText(item.Line + "   " + item.ItemCode + "   " + item.ItemDesc + "    \n");
                    foreach (var item2 in item.Children)
                    {
                        textBox1.AppendText("L   " + item2.ItemCode + "    " + item2.ItemDesc + "   " + item2.Quantity + "\n");
                    }
                    int i = 0;
                    foreach (var item2 in item.Children)
                    {
                    
                        var child = children.Find(x => x.ItemCode == item2.ItemCode);
                    if (child.Children.Count > 0)
                    {
                        textBox1.AppendText(child.Line + "   " + child.ItemCode + "   " + child.ItemDesc + "    \n");
                        foreach (var child2 in child.Children)
                        {
                            textBox1.AppendText("L   " + child2.ItemCode + "    " + child2.ItemDesc + "   " + child2.Quantity + "\n");
                        }
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

