﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelHelper
{
    /// <summary>
    /// SourceAnalysis
    /// </summary>
    public class Source
    {
        public SourceType sourceType { get; set; }
        public string sourceName { get; set; }
        public string oracleTableName { get; set; }
        public List<SourceCol> sourceCols { get; set; }
        public StopType stopType { get; set; }
        public string stopKey { get; set; }
        public string colKey { get; set; }
        public string[] NoCols { get; set; }
        public string[] GuidCols { get; set; }
        public int firstRowNo {
            get
            {
                if (sourceCols.Count == 0)
                {
                    return 0;
                }
                return sourceCols.First().rowIndex;
            }
        }
        public int firstNotFixedRowNo
        {
            get
            {
                if (sourceCols.Count == 0)
                {
                    return 0;
                }
                return sourceCols.First(x => !x.isFixed).rowIndex;
            }
        }
        public int firstColNo
        {
            get
            {
                if (sourceCols.Count == 0)
                {
                    return 0;
                }
                return sourceCols.First().colIndex;
            }
        }

        public int firstNotFixedColNo
        {
            get
            {
                if (sourceCols.Count == 0)
                {
                    return 0;
                }
                return sourceCols.First(x => !x.isFixed).colIndex;
            }
        }
        public int[] rowIndexRange
        {
            get
            {
                List<int> result = new List<int>() { };
                foreach (SourceCol col in sourceCols)
                {
                    result.Add(col.rowIndex);
                }
                if (result.Count == 0)
                {
                    return result.ToArray();
                }
                List<int> newResult = new List<int>() { };
                for (int i = result.Min(); i <= result.Max(); i++)
                {
                    newResult.Add(i);
                } 
                return newResult.ToArray();
            }
        }
        public int[] rowIndexCollection
        {
            get
            {
                List<int> result = new List<int>() { };
                foreach (SourceCol col in sourceCols)
                {
                    if (!col.isPublic)
                    {
                        result.Add(col.rowIndex);
                    }
                }
                return result.ToArray();
            }
        }
        public int[] rowIndexFullCollection
        {
            get
            {
                List<int> result = new List<int>() { };
                foreach (SourceCol col in sourceCols)
                {
                    result.Add(col.rowIndex);
                }
                return result.ToArray();
            }
        }
        public int[] colIndexRange
        {
            get
            {
                List<int> result = new List<int>() { };
                foreach (SourceCol col in sourceCols)
                {
                    result.Add(col.colIndex);
                }
                if (result.Count == 0)
                {
                    return result.ToArray();
                }
                List<int> newResult = new List<int>() { };
                for (int i = result.Min(); i <= result.Max(); i++)
                {
                    newResult.Add(i);
                }
                return newResult.ToArray();
            }
        }
        public int[] colIndexCollection
        {
            get
            {
                List<int> result = new List<int>() { };
                foreach (SourceCol col in sourceCols)
                {
                    if (!col.isPublic)
                    {
                        result.Add(col.colIndex);
                    }
                }
                return result.ToArray();
            }
        }
        public int[] colIndexFullCollection
        {
            get
            {
                List<int> result = new List<int>() { };
                foreach (SourceCol col in sourceCols)
                {
                   result.Add(col.colIndex);
                }
                return result.ToArray();
            }
        }

    }
    public class SourceCol
    {
        public string colName { get; set; }
        public string position { get; set; }
        public int rowIndex
        {
            get
            {
                int index = position.IndexOf("-");
                return Convert.ToInt32(position.Substring(0, index));
            }
        }
        public int colIndex
        {
            get
            {
                int index = position.IndexOf("-");
                return Convert.ToInt32(position.Substring(index + 1, position.Length - index - 1));
            }
        }
        public bool isPublic { get; set; }
        public bool isFixed { get; set; }
        public bool isDate { get; set; }
        public string dateFormat { get; set; }
    }
    public enum SourceType
    {
        Common = 0,
        Down,
        Right
    }
    public enum StopType
    {
        StopWhileKeyEmpty = 0,
        StopWhileHasEmpty,
        StopWhileAllEmpty,
        StopWithWord,
    }
}
