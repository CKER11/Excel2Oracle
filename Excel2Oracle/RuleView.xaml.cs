﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using ExcelHelper;
using System.Data;
using System.IO;
namespace Excel2Oracle
{
    /// <summary>
    /// RuleView.xaml 的交互逻辑
    /// </summary>
    public partial class RuleView : Window
    {
       
        public RuleView()
        {
            InitializeComponent();
        }

        public RuleView(List<Source> result)
        {
            InitializeComponent();
            GetCurRuleIntroduction();
            GetCurRule(result);
        }

        private void GetCurRuleIntroduction()
        {
            string exeDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string introductionPath = Path.Combine(exeDirectory, "RuleHelper.txt");
            if (File.Exists(introductionPath))
                lblRulesIntroduction.Content = File.ReadAllText(introductionPath);
        }

        private void GetCurRule(List<Source> result)
        {
            if (result == null)
            {
                return;
            }
            DataTable dt = new DataTable();
            dt.Columns.Add("序号");
            dt.Columns.Add("数据源表名");
            dt.Columns.Add("Oracle表名");
            dt.Columns.Add("主键");
            dt.Columns.Add("读取类型");
            dt.Columns.Add("开始读取单元格");
            dt.Columns.Add("停止读取类型");
            dt.Columns.Add("列数量");
            dt.Columns.Add("列名");
            int index = 0;
            foreach (Source source in result)
            {
                index++;
                DataRow dr = dt.NewRow();
                dr[0] = index;
                dr[1] = source.sourceName;
                dr[2] = source.oracleTableName;
                dr[3] = source.colKey;
                dr[4] = GetStrType(source.sourceType);
                dr[5] = GetStrStartCols(source.sourceCols);
                dr[6] = GetStrStopType(source.sourceType, source.stopType, source.stopKey, source.colKey);
                dr[7] = source.sourceCols.Count;
                dr[8] = GetStrSourceCols(source.sourceCols);
                dt.Rows.Add(dr);
            }
            lvResultView.ItemsSource = dt.DefaultView;
        }

        private object GetStrStartCols(List<SourceCol> sourceCols)
        {
            StringBuilder str = new StringBuilder();
            foreach (SourceCol col in sourceCols)
            {
                str.Append(AnalysisHelper.GetPositionByKey(col.position) + ";");
            }
            return str.ToString();
        }

        private object GetStrStopType(SourceType sourceType,StopType stopType, string stopKey,string colKey)
        {
            if (sourceType== SourceType.Common)
            {
                return "一一对应";
            }
            string str = "";
            switch (stopType)
            {
                case StopType.StopWhileHasEmpty:
                    str = "读到有空字符的行即结束";
                    break;
                case StopType.StopWhileKeyEmpty:
                    str = "读到列\""+ colKey + "\"为空字符即结束";
                    break;
                case StopType.StopWhileAllEmpty:
                    str = "读到全为空字符的行即结束";
                    break;
                case StopType.StopWithWord:
                    str = "读到关键词\"" + stopKey + "\"即结束";
                    break;
                default:
                    break;
            }
            return str;
        }

        private object GetStrSourceCols(List<SourceCol> sourceCols)
        {
            StringBuilder str = new StringBuilder();
            foreach (SourceCol col in sourceCols)
            {
                str.Append(col.colName+";");
            }
            return str.ToString();
        }

        private string GetStrType(SourceType sourceType)
        {
            string str = "";
            switch (sourceType)
            {
                case SourceType.Common:
                    str = "普通对应";
                    break;
                case SourceType.Down:
                    str = "向下自增";
                    break;
                case SourceType.Right:
                    str = "向右自增";
                    break;
                default:
                    break;
            }
            return str;
        }
    }
}
