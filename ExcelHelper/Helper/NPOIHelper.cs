﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Data;
using System.Windows;
using NPOI.SS.Util;

namespace ExcelHelper
{
    public class ExcelHelper : IDisposable
    {
        private string fileName = null; //文件名
        private IWorkbook workbook = null;
        private FileStream fs = null;
        private bool disposed;

        public ExcelHelper(string fileName)
        {
            this.fileName = fileName;
            disposed = false;
        }

        internal Dictionary<string, string> ReadAllContentText(string sheetName = null)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                ISheet sheet = null;
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.Trim().IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.Trim().IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);
                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    if (sheet == null)
                        sheet = workbook.GetSheetAt(0);
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow headerRow = sheet.GetRow(0);
                    //最后一个方格的编号 即总的列数
                    int colNum = headerRow.LastCellNum;
                    //最后一行的标号  即总的行数
                    int rowNum = sheet.LastRowNum;
                    for (int i = 0; i <= rowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null)
                        {
                            continue;
                        }
                        int cellCount = row.Cells.Count;
                        for (int j = 0; j < cellCount; j++)
                        {
                            ICell cell = row.Cells[j];
                            string value = GetValueByType(cell, cell.CellType);
                            if (cell.CellType != CellType.Blank && value.Trim() != string.Empty)
                            {
                                string key = string.Format("{0}-{1}", cell.RowIndex + 1, cell.ColumnIndex + 1);
                                dic[key] = value.Trim();
                            }
                        }
                    }

                }
                return dic;
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Exception: " + ex.Message);
                MessageBox.Show("Exception: " + ex.Message);
                return dic;
            }
        }

        private string GetValueByType(ICell cell, CellType cellType)
        {
            try
            {
                switch (cellType)
                {
                    case CellType.Unknown:
                        return cell.StringCellValue;
                    case CellType.Numeric:
                        if (HSSFDateUtil.IsCellDateFormatted(cell) && cell.DateCellValue != null)
                            return cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss");
                        return cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Formula:
                        return cell.NumericCellValue.ToString(); ;
                    case CellType.Blank:
                        return string.Empty;
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Error:
                        return cell.ErrorCellValue.ToString();
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(cell.RowIndex + "行" + cell.ColumnIndex + "列");
                return "读取错误";
            }
        }

        /// <summary>
        /// 数字转字母，用于excel列
        /// </summary>
        /// <param name="index"></param>
        /// <param name="isLower"></param>
        /// <returns></returns>
        public static string NumToAlpha(int index, bool isLower = true)
        {
            index--;
            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)(isLower ? 'a' : 'A'))).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);
            return String.Join(string.Empty, chars.ToArray());
        }
        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;

            fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                workbook = new XSSFWorkbook();
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
                workbook = new HSSFWorkbook();

            try
            {
                if (workbook != null)
                {
                    sheet = workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }

                if (isColumnWritten == true) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                workbook.Write(fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);

                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    {
                        sheet = workbook.GetSheetAt(0);
                    }
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow firstRow = sheet.GetRow(0);
                    int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                    if (isFirstRowColumn)
                    {
                        for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                string cellValue = cell.StringCellValue;
                                if (cellValue != null)
                                {
                                    DataColumn column = new DataColumn(cellValue);
                                    data.Columns.Add(column);
                                }
                            }
                        }
                        startRow = sheet.FirstRowNum + 1;
                    }
                    else
                    {
                        startRow = sheet.FirstRowNum;
                    }

                    //最后一列的标号
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; ++i)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue; //没有数据的行默认是null　　　　　　　

                        DataRow dataRow = data.NewRow();
                        for (int j = row.FirstCellNum; j < cellCount; ++j)
                        {
                            if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                dataRow[j] = row.GetCell(j).ToString();
                        }
                        data.Rows.Add(dataRow);
                    }
                }

                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    if (fs != null)
                        fs.Close();
                }

                fs = null;
                disposed = true;
            }
        }

        internal Dictionary<string, string> ReadAllContent(string sheetName = null)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                ISheet sheet = null;
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                if (fileName.Trim().IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook(fs);
                else if (fileName.Trim().IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook(fs);
                if (sheetName != null)
                {
                    sheet = workbook.GetSheet(sheetName);
                    //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    if (sheet == null)
                        sheet = workbook.GetSheetAt(0);
                }
                else
                {
                    sheet = workbook.GetSheetAt(0);
                }
                if (sheet != null)
                {
                    IRow headerRow = sheet.GetRow(0);
                    //最后一个方格的编号 即总的列数
                    int colNum = headerRow.LastCellNum;
                    //最后一行的标号  即总的行数
                    int rowNum = sheet.LastRowNum;
                    //获得NPOI提供的合并单元信息
                    List<CellRangeAddress> MergedCellAdInfo = new List<CellRangeAddress>();
                    for (int i = 0; i < sheet.NumMergedRegions; i++)
                    {
                        MergedCellAdInfo.Add(sheet.GetMergedRegion(i));
                    }
                    //获得所有的值
                    for (int i = 0; i <= rowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        if (row == null)
                        {
                            continue;
                        }
                        int cellCount = row.Cells.Count;
                        for (int j = 0; j < cellCount; j++)
                        {
                            ICell cell = row.Cells[j];
                            string value = GetValueByType(cell, cell.CellType);
                            string key = string.Format("{0}-{1}", cell.RowIndex + 1, cell.ColumnIndex + 1);
                            dic[key] = value.Trim();
                        }
                    }
                    var keys = dic.Keys.ToList();
                    //根据NPOI提供的合并单元信息，将合并单元格的值设为首单元的内容
                    for (int i = 0; i < dic.Count; i++)
                    {
                        string key = dic.Keys.ToList()[i];
                        int[] pos = AnalysisHelper.GetPositionIndexByKey(key);
                        int index = MergedCellAdInfo.FindIndex(x => x.IsInRange(pos[0] - 1, pos[1] - 1));
                        if (index >= 0)
                        {
                            var adInfo = MergedCellAdInfo[index];
                            string adKey = string.Format("{0}-{1}", adInfo.FirstRow + 1, adInfo.FirstColumn + 1);
                            dic[key] = dic[adKey];
                        }
                    }
                }
                return dic;
            }
            catch (Exception ex)
            {
                if (ex.Message.IndexOf("指定的参数已超出有效值的范围。")>=0)
                {
                    MessageBox.Show("Exception: Excel打开失败，如果是WPS保存的Excel文件，请使用Microsoft Office Excel重新保存");
                    return dic;
                }
                MessageBox.Show("Exception: " + ex.Message);
                return dic;
            }
        }
    }
}