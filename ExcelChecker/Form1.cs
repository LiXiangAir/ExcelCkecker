using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelChecker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void browse_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "xls文件|*.xls|xlsx文件|*.xlsx|所有文件|*.*";
            openFileDialog1.RestoreDirectory = true;

            var dialogResult = openFileDialog1.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                importExcel.Enabled = true;
            }                
        }

        private void importExcel_Click(object sender, EventArgs e)
        {
            var filePath = textBox1.Text;
            if (filePath == null || filePath.Equals(""))
            {
                MessageBox.Show(null, "未选择Excel文件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            checkExcel(filePath);
        }

        private void checkExcel(string filePath)
        {
            try
            {
                HashSet<string> set = new HashSet<string>();
                FileStream frs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
                HSSFWorkbook wb = new HSSFWorkbook(frs);
                frs.Close();

                var whiteStyle = wb.CreateCellStyle();
                whiteStyle.FillForegroundColor = HSSFColor.White.Index;
                whiteStyle.FillPattern = FillPattern.SolidForeground;
                whiteStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                whiteStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                whiteStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                whiteStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;

                var yellowStyle = wb.CreateCellStyle();
                yellowStyle.FillForegroundColor = HSSFColor.Yellow.Index;
                yellowStyle.FillPattern = FillPattern.SolidForeground;
                yellowStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                yellowStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                yellowStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                yellowStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                
                int count = 0;
                for (int i = 0; i < wb.NumberOfSheets; i++)
                {
                    var sheet = wb.GetSheetAt(i);
                    if (sheet == null) continue;

                    for (int j = sheet.FirstRowNum; j <= sheet.LastRowNum; j++)
                    {
                        var row = sheet.GetRow(j);
                        if (row == null) continue;

                        for (int k = row.FirstCellNum; k < row.LastCellNum; k++)
                        {
                            var cell = row.GetCell(k);
                            if (cell == null) continue;

                            var word = cell.ToString();
                            if (word == null || word.Equals("")) continue;

                            cell.CellStyle = whiteStyle;

                            if (!set.Contains(word))
                            {
                                set.Add(word);
                            }
                            else
                            {                              
                                cell.CellStyle = yellowStyle;                                
                                count++;
                            }
                        }
                    }
                }

                FileStream fws = System.IO.File.OpenWrite(filePath);
                wb.Write(fws);
                fws.Close(); 
  
                MessageBox.Show(null, "共有" + count + "个重复词汇，已在文件中标黄", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(null, e.ToString(), "完成", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }            
        }        
    }
}
