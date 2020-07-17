using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HSSF.Model;
using NPOI.HSSF.Record.Chart;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace excelMerge
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void SelectFile1_Click(object sender, EventArgs e)
        {
            file1.Text = SelectFile();
        }

        private string SelectFile()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "请选择Excel文件|*.xls;*.xlsx";
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return dialog.FileName;
            }

            return "";
        }

        private void SelectFile2_Click(object sender, EventArgs e)
        {
            file2.Text = SelectFile();
        }

        private void SelectFile3_Click(object sender, EventArgs e)
        {
            file3.Text = SelectFile();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (CheckSelectFile() == false)
            {
                return;
            }

            // 创建即将导出的excel
            HSSFWorkbook newWorkBook = new HSSFWorkbook();
            ISheet newSheet = newWorkBook.CreateSheet("sheet test");
            IRow firstRow = newSheet.CreateRow(0);

            // 读取第一个excel文件
            IWorkbook fileOneWorkbook = WorkbookFactory.Create(file1.Text);
            ISheet fileOneSheet = fileOneWorkbook.GetSheetAt(0);

            // 读取第二个excel文件
            IWorkbook fileTwoWorkbook = WorkbookFactory.Create(file2.Text);
            ISheet fileTwoSheet = fileTwoWorkbook.GetSheetAt(0);

            // 读取第三个excel文件
            IWorkbook fileThreeWorkbook = WorkbookFactory.Create(file3.Text);
            ISheet fileThreeSheet = fileThreeWorkbook.GetSheetAt(0);

            // 复制表头
            IRow fileOneRow = fileOneSheet.GetRow(0);
            CopyRow(fileOneRow, firstRow, false);

            // 拼接第二个excel和第三个excel的部分表头
            formatTableTitle(firstRow, fileTwoSheet, fileThreeSheet);

            for (int i = 1; i <= fileOneSheet.LastRowNum; i++)
            {
                fileOneRow = fileOneSheet.GetRow(i);
                if (fileOneRow == null)
                {
                    continue;
                }

                IRow newRow = newSheet.CreateRow(i);
                CopyRow(fileOneRow, newRow);

                string date = newRow.GetCell(0).StringCellValue;
                string sku = newRow.GetCell(4).StringCellValue;
               
                HandleTwoFile(sku, date, newRow, fileTwoSheet);
                HandleThreeFile(sku, date, newRow, fileThreeSheet);
            }

            // 将excel存储到用户指定路径
            SaveFileDialog savedialog = new SaveFileDialog();
            savedialog.Filter = "Excel 2007|*.xls|Excel 2013|*.xlsx";
            savedialog.FilterIndex = 0;
            savedialog.CheckPathExists = true;
            if (savedialog.ShowDialog() == DialogResult.OK)
            {
                FileStream newExcel = File.OpenWrite(savedialog.FileName);
                newWorkBook.Write(newExcel);
                MessageBox.Show("合并完成！", "提示");

                newExcel.Close();
            }

            // 释放资源
            newWorkBook.Close();
            fileOneWorkbook.Close();
            fileTwoWorkbook.Close();
            fileThreeWorkbook.Close();
        }

        private void formatTableTitle(IRow firstRow, ISheet fileTwoSheet, ISheet fileThreeSheet)
        {
            IRow twoTableTitleRow = fileTwoSheet.GetRow(0);
            string quantityTitle = twoTableTitleRow.GetCell(14).StringCellValue;
            string itemPromotionDiscountTitle = twoTableTitleRow.GetCell(22).StringCellValue;

            IRow threeTableTitleRow = fileThreeSheet.GetRow(0);
            string clickTitle = threeTableTitleRow.GetCell(8).StringCellValue;
            string totalOneTitle = threeTableTitleRow.GetCell(19).StringCellValue;
            string totalTwoTitle = threeTableTitleRow.GetCell(18).StringCellValue;
            string moneyTitle = threeTableTitleRow.GetCell(11).StringCellValue;

            ICell newTitleCell;
            newTitleCell = firstRow.CreateCell(firstRow.LastCellNum);
            newTitleCell.SetCellValue(itemPromotionDiscountTitle);

            newTitleCell = firstRow.CreateCell(firstRow.LastCellNum);
            newTitleCell.SetCellValue(quantityTitle);

            newTitleCell = firstRow.CreateCell(firstRow.LastCellNum);
            newTitleCell.SetCellValue(clickTitle);

            newTitleCell = firstRow.CreateCell(firstRow.LastCellNum);
            newTitleCell.SetCellValue(totalTwoTitle);

            newTitleCell = firstRow.CreateCell(firstRow.LastCellNum);
            newTitleCell.SetCellValue(totalOneTitle);
                     
            newTitleCell = firstRow.CreateCell(firstRow.LastCellNum);
            newTitleCell.SetCellValue(moneyTitle);
        }

        private void HandleThreeFile(string sku, string date, IRow newRow, ISheet fileThreeSheet)
        {
            double money = 0;
            int click = 0, totalOne = 0, totalTwo = 0;
            ICell moneyCell, clickCell, totalOneCell, totalTwoCell;
            IRow fileRow;
            ICell newCell;
            for (int i = 1; i <= fileThreeSheet.LastRowNum; i++)
            {
                fileRow = fileThreeSheet.GetRow(i);
                if (fileRow == null)
                {
                    continue;
                }

                string rowSku = fileRow.GetCell(5).ToString();
                string rowDate = formatDate(fileRow.GetCell(0));

                if (rowSku == sku && rowDate == date)
                {
                    clickCell = fileRow.GetCell(8);
                    if (clickCell != null && clickCell.CellType == CellType.Numeric)
                    {
                        click += (int)clickCell.NumericCellValue;
                    }

                    totalOneCell = fileRow.GetCell(19);
                    if (totalOneCell != null && totalOneCell.CellType == CellType.Numeric)
                    {
                        totalOne += (int)fileRow.GetCell(19).NumericCellValue;
                    }

                    totalTwoCell = fileRow.GetCell(18);
                    if (totalTwoCell != null && totalTwoCell.CellType == CellType.Numeric)
                    {
                        totalTwo += (int)fileRow.GetCell(18).NumericCellValue;
                    }

                    moneyCell = fileRow.GetCell(11);
                    if (moneyCell != null && moneyCell.CellType == CellType.Numeric)
                    {
                        money += moneyCell.NumericCellValue;
                    }
                }
            }

            newCell = newRow.CreateCell(newRow.LastCellNum);
            newCell.SetCellValue(click);

            newCell = newRow.CreateCell(newRow.LastCellNum);
            newCell.SetCellValue(totalTwo);

            newCell = newRow.CreateCell(newRow.LastCellNum);
            newCell.SetCellValue(totalOne);
                       
            newCell = newRow.CreateCell(newRow.LastCellNum);
            newCell.SetCellValue(money);
        }

        private void HandleTwoFile(string sku, string date, IRow newRow, ISheet fileTwoSheet)
        {
            int quantity = 0, itemPromotionDiscount = 0;
            ICell quantityCell, itemPromotionDiscountCell;
            IRow fileRow;
            ICell newCell;
            for (int i = 1; i <= fileTwoSheet.LastRowNum; i++)
            {
                fileRow = fileTwoSheet.GetRow(i);
                if (fileRow == null)
                {
                    continue;
                }

                string rowSku = fileRow.GetCell(11).ToString();
                string rowDate = formatDate(fileRow.GetCell(2));

                if (rowSku == sku && rowDate == date)
                {
                    quantityCell = fileRow.GetCell(14);
                    if (quantityCell != null && quantityCell.CellType == CellType.Numeric)
                    {
                        quantity += (int)quantityCell.NumericCellValue;
                    }

                    itemPromotionDiscountCell = fileRow.GetCell(22);
                    if (itemPromotionDiscountCell != null && itemPromotionDiscountCell.CellType == CellType.Numeric)
                    {
                        itemPromotionDiscount += (int)itemPromotionDiscountCell.NumericCellValue;
                    }
                }
            }

            newCell = newRow.CreateCell(newRow.LastCellNum);
            newCell.SetCellValue(itemPromotionDiscount);

            newCell = newRow.CreateCell(newRow.LastCellNum);
            newCell.SetCellValue(quantity);
        }

        private string formatDate(ICell dateCell)
        {
            string dateStr;
            if (dateCell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(dateCell))
            {
                dateStr = dateCell.DateCellValue.ToString();
            }
            else
            {
                dateStr = dateCell.ToString();
            }

            DateTime dateA = DateTime.Parse(dateStr);

            return dateA.ToString();
        }

        private void CopyRow(IRow fileRow, IRow newRow, bool notTableTitle = true)
        {
            ICell newCell, fileOneCell;
            for (int j = 0; j < fileRow.LastCellNum; j++)
            {
                newCell = newRow.CreateCell(j);
                fileOneCell = fileRow.GetCell(j);

                if (j == 0 && notTableTitle)
                {
                    newCell.SetCellValue(formatDate(fileOneCell));
                }
                else
                {
                    newCell.SetCellValue(fileOneCell.ToString());
                }              
            }
        }

        private bool CheckSelectFile()
        {
            if (file1.Text == "" || file2.Text == "" || file3.Text == "")
            {
                MessageBox.Show("请确保Excel文件已全部选择", "操作错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (File.Exists(file1.Text) == false)
            {
                MessageBox.Show("文件1不存在", "操作错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (File.Exists(file2.Text) == false)
            {
                MessageBox.Show("文件2不存在", "操作错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (File.Exists(file3.Text) == false)
            {
                MessageBox.Show("文件3不存在", "操作错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

    }
}
