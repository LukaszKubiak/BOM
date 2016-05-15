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
        Excel.Application importApp;
        Excel.Workbook importWorkBook;
        Excel.Worksheet importWorkSheet;

        public Form1()
        {
            InitializeComponent();
            importApp = new Excel.Application();
            importWorkBook = importApp.Workbooks.Open(@"C:\Users\Lukasz\Documents\Import.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            importWorkSheet = (Excel.Worksheet)importWorkBook.Worksheets.get_Item(1);
        }

        private void Open_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                CalculateSheet();
            }
        }

        private void insertData(Excel.Worksheet sheet, int row, string line, string code, string desc, string quantity)
        {
            sheet.Cells[row, 3] = line;
            sheet.Cells[row, 5] = code;
            sheet.Cells[row, 6] = desc;
            sheet.Cells[row, 12] = quantity;
        }
        private void CalculateSheet()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(openFileDialog1.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); ;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range range = xlWorkSheet.UsedRange;
            object misValue = System.Reflection.Missing.Value;
            var ImportSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Add(Type.Missing, xlWorkSheet, Type.Missing, Type.Missing);
            importWorkSheet.UsedRange.Copy();
            ImportSheet.Paste();
            int rCnt = 0;
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
                if (childCode != "" || childDesc != "")
                    item.Children.Add(new Child(childCode, childDesc, childQuantity));
                int quantityIndex = rCnt;
                while ((range.Cells[++rCnt, 2] as Excel.Range).Text == "" && rCnt <= range.Rows.Count)
                {
                    childCode = (range.Cells[rCnt, 3] as Excel.Range).Text;
                    childDesc = (range.Cells[rCnt, 4] as Excel.Range).Text;
                    childQuantity = (range.Cells[rCnt, 7] as Excel.Range).Text;
                    if (childCode == "" && childDesc == "" && !(range.Cells[rCnt, 3] as Excel.Range).MergeCells)
                    {
                        childCode = (range.Cells[rCnt, 11] as Excel.Range).Text;
                        childDesc = (range.Cells[rCnt, 12] as Excel.Range).Text;
                        childQuantity = (range.Cells[rCnt, 14] as Excel.Range).Text;
                    }
                    if (childQuantity == "" && (range.Cells[rCnt, 7] as Excel.Range).MergeCells)
                        childQuantity = (range.Cells[quantityIndex, 7] as Excel.Range).Text;
                    else
                        quantityIndex = rCnt;
                    if (childCode != "" || childDesc != "")
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
                    for (rCnt = 18; rCnt <= range.Rows.Count; rCnt++)
                    {
                        if ((range.Cells[rCnt, 3] as Excel.Range).Text == item.ItemCode)
                        {
                            string childCode = (range.Cells[rCnt, 11] as Excel.Range).Text;
                            string childDesc = (range.Cells[rCnt, 12] as Excel.Range).Text;
                            string childQuantity = (range.Cells[rCnt, 14] as Excel.Range).Text;
                            if ((childCode != "" || childDesc != "") && childCode != item.ItemCode)
                                item.Children.Add(new Child(childCode, childDesc, childQuantity));
                            while ((range.Cells[++rCnt, 3] as Excel.Range).Text == "" && rCnt <= range.Rows.Count && (range.Cells[rCnt, 3] as Excel.Range).MergeCells)
                            {
                                childCode = (range.Cells[rCnt, 11] as Excel.Range).Text;
                                childDesc = (range.Cells[rCnt, 12] as Excel.Range).Text;
                                childQuantity = (range.Cells[rCnt, 14] as Excel.Range).Text;
                                if (childCode != "" || childDesc != "")
                                    item.Children.Add(new Child(childCode, childDesc, childQuantity));
                            }
                            break;
                        }
                    }
                    children.Add(item);
                }

            }
            rCnt = 8;
            foreach (var item in items)
            {

                textBox1.AppendText(item.Line + "   " + item.ItemCode + "   " + item.ItemDesc + "    \n");
                insertData(ImportSheet, rCnt++, item.Line, item.ItemCode, item.ItemDesc, "");
                foreach (var item2 in item.Children)
                {
                    textBox1.AppendText("L   " + item2.ItemCode + "    " + item2.ItemDesc + "   " + item2.Quantity + "\n");
                    insertData(ImportSheet, rCnt++, "L", item2.ItemCode, item2.ItemDesc, item2.Quantity);
                }
                int i = 0;
                foreach (var item2 in item.Children)
                {

                    var child = children.Find(x => x.ItemCode == item2.ItemCode);
                    if (child.Children.Count > 0)
                    {
                        textBox1.AppendText(child.Line + "   " + child.ItemCode + "   " + child.ItemDesc + "    \n");
                        insertData(ImportSheet, rCnt++, child.Line, child.ItemCode, child.ItemDesc, "");
                        foreach (var child2 in child.Children)
                        {
                            textBox1.AppendText("L   " + child2.ItemCode + "    " + child2.ItemDesc + "   " + child2.Quantity + "\n");
                            insertData(ImportSheet, rCnt++, "L", child2.ItemCode, child2.ItemDesc, child2.Quantity);
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

