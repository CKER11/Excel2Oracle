﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;

namespace ExcelHelper
{
    public partial class AnalysisHelper
    {
        public static bool Analysis(string templateFilePath, ProcessReport report, ProcessReportClear reportClear, out List<Source> result, out Dictionary<string, string> contentConfig, out Dictionary<string, string[]> guidInfo, out Dictionary<string, string[]> noInfo, out Dictionary<string, string> beforeSql, out Dictionary<string, string> afterSql)
        {
            beforeSql = new Dictionary<string, string>();
            afterSql = new Dictionary<string, string>();
            guidInfo = new Dictionary<string, string[]>();
            noInfo = new Dictionary<string, string[]>();
            result = new List<Source>() { };
            contentConfig = new Dictionary<string, string>();
            reportClear();
            report("开始读取模板文件");
            ExcelHelper excelHelper = new ExcelHelper(templateFilePath);
            report("开始读取所有内容");
            //读取所有有文本的单元格并保存
            Dictionary<string, string> AllContent = excelHelper.ReadAllContentText();

            if (AllContent.Count == 0)
            {
                report("解析失败，请修改配置后重试");
                return false;
            }
            report("开始读取数据源");
            if (!AllContent.Keys.Contains("1-1"))
            {
                MessageBox.Show("未读取到数据源，请在A1结尾处配置数据源，如：\r\n标题(ds1=table;ds2=table2)");
                report("未读取到数据源，请在A1结尾处配置数据源，如：\r\n标题(ds1=table;ds2=table2)");
                return false;
            }
            //忽略括弧、分号、逗号的中英文区别
            string strDBSource = AllContent["1-1"].Replace("【", "[").Replace("】", "]").Replace("（", "(").Replace("）", ")").Replace("，", ",").Replace("；", ";"); ;
            //数据源最后是否包含括弧或者中括号的配置
            Regex reg = new Regex(@"\((?<=\().+(?<=\))$");
            Match match = reg.Match(strDBSource);
            if (!match.Success)
            {
                MessageBox.Show("数据源配置错误，请在A1结尾处配置数据源，如：\r\n标题(ds1=table;ds2=table2)");
                report("未读取到数据源，请在A1结尾处配置数据源，如：\r\n标题(ds1=table;ds2=table2)");
                return false;
            }
            string strDBConfig = "";
            if (match.Success)
            {
                strDBConfig = match.Value;
            }
            if (string.IsNullOrEmpty(strDBConfig))
            {
                MessageBox.Show("数据源配置错误，请在A1结尾处配置数据源，如：\r\n标题(ds1=table;ds2=table2)");
                report("未读取到数据源，请在A1结尾处配置数据源，如：\r\n标题(ds1=table;ds2=table2)");
                return false;
            }
            Dictionary<string, string> DBSource = new Dictionary<string, string>();
            string strDB = strDBConfig.Substring(1, strDBConfig.Length - 2);
            if (strDB.EndsWith(";"))
                strDB = strDB.Remove(strDB.Length - 1);
            string regConfigs = @"\[[^\[\]]+\]";
            //读取数据源配置项
            string strDBTemp = strDB;
            MatchCollection cfgs = new Regex(regConfigs).Matches(strDB);
            foreach (Match item in cfgs)
            {
                strDBTemp = strDBTemp.Replace(item.Value, " ".PadLeft(item.Length, ' '));
            }

            try
            {
                //解析数据源
                string[] strDBs = strDBTemp.Replace(" ", "").Split(';');
                foreach (string oneDB in strDBs)
                {
                    int eqIndex = oneDB.IndexOf("=");
                    string dbKey = oneDB.Substring(0, eqIndex);
                    string dbTableName = oneDB.Substring(eqIndex + 1, oneDB.Length - eqIndex - 1);
                    DBSource.Add(dbKey, dbTableName);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("数据源格式错误\r\n" + e.Message);
                report("数据源格式错误\r\n" + e.Message);
                return false;
            }
            if (!strDBTemp.EndsWith(";"))
                strDBTemp += ";";
            List<int[]> DBSourceIndex = new List<int[]>();
            int sIndex = 0;
            int eIndex = 0;
            string str = strDBTemp;
            while (str.IndexOf(";", sIndex) >= 0)
            {
                int splitIndex = str.IndexOf(";", sIndex);
                eIndex = splitIndex;
                DBSourceIndex.Add(new int[] { sIndex, eIndex });
                sIndex = splitIndex + 1;
            }
            if (DBSourceIndex.Count != DBSource.Count)
            {
                MessageBox.Show("数据源格式错误：获取的索引数量与数据源数量不匹配");
                report("数据源格式错误：获取的索引数量与数据源数量不匹配");
                return false;
            }
            for (int k = 0; k < DBSourceIndex.Count; k++)
            {
                int[] item = DBSourceIndex[k];
                string dbKey = DBSource.Keys.ToList()[k];
                for (int i = 0; i < cfgs.Count; i++)
                {
                    Match m = cfgs[i];
                    if (m.Index > item[0] && m.Index < item[1])
                    {
                        //存在beforeSql
                        if (m.Value.ToLower().IndexOf("beforesql") >= 0)
                        {
                            string strConfig = m.Value;
                            strConfig = strConfig.Substring(1, strConfig.Length - 2);
                            int eqIndex = strConfig.IndexOf("=");
                            Match bm = new Regex(@"\s*beforesql\s*=", RegexOptions.IgnoreCase).Match(strConfig);
                            string sql = strConfig.Substring(bm.Index + bm.Length, strConfig.Length - bm.Index - bm.Length);
                            beforeSql.Add(dbKey, sql);
                        }
                        //存在AfterSql
                        else if (m.Value.ToLower().IndexOf("aftersql") >= 0)
                        {
                            string strConfig = m.Value;
                            strConfig = strConfig.Substring(1, strConfig.Length - 2);
                            int eqIndex = strConfig.IndexOf("=");
                            Match bm = new Regex(@"\s*aftersql\s*=", RegexOptions.IgnoreCase).Match(strConfig);
                            string sql = strConfig.Substring(bm.Index + bm.Length, strConfig.Length - bm.Index - bm.Length);
                            afterSql.Add(dbKey, sql);
                        }
                        //存在Guid列或No列
                        else if (m.Value.ToLower().IndexOf("no") >= 0 || m.Value.ToLower().IndexOf("guid") >= 0)
                        {
                            string strConfig = m.Value.Replace(" ", "");
                            strConfig = strConfig.Substring(1, strConfig.Length - 2);
                            string[] strCols = strConfig.Split(',');
                            foreach (string col in strCols)
                            {
                                int eqIndex2 = col.IndexOf("=");
                                string colKey = col.Substring(0, eqIndex2);
                                string colvalue = col.Substring(eqIndex2 + 1, col.Length - eqIndex2 - 1);
                                if (colKey.ToLower() == "no")
                                {
                                    noInfo.Add(dbKey, colvalue.Split('&'));
                                }
                                if (colKey.ToLower() == "guid")
                                {
                                    guidInfo.Add(dbKey, colvalue.Split('&'));
                                }
                            }
                        }
                    }
                }
            }
            if (DBSource.Count == 0)
            {
                MessageBox.Show("未找到任何数据源");
                report("未找到任何数据源");
                return false;
            }

            report("开始读取所有配置");
            //找到所有含有数据源名称(如ds1、ds2)的配置
            foreach (string key in DBSource.Keys)
            {
                Dictionary<string, string> keyConfig = AllContent.Where(x => x.Value.Trim().IndexOf(key + ".") >= 0).ToDictionary(x => x.Key, y => y.Value);
                foreach (var item in keyConfig)
                {
                    if (!contentConfig.Keys.Contains(item.Key))
                    {
                        contentConfig.Add(item.Key, item.Value);
                    }
                }
            }
            Dictionary<string, string> dbConfig = new Dictionary<string, string>();
            //去除A1数据源配置
            contentConfig.Remove("1-1");
            //检测警告
            StringBuilder strWarnBuilder = new StringBuilder();
            foreach (var item in contentConfig)
            {
                if (!VaildConfig(DBSource.Keys.ToList(), item.Value))
                {
                    strWarnBuilder.AppendFormat("警告：不识别的配置：{0}处的\"{1}\"\r\n", AnalysisHelper.GetPositionByKey(item.Key), item.Value);
                }
            }
            if (!string.IsNullOrEmpty(strWarnBuilder.ToString()))
                report(strWarnBuilder.ToString());
            //去除DownEnd和RightEnd
            dbConfig = contentConfig.Where(x => x.Value.ToLower().IndexOf("downend") < 0 && x.Value.ToLower().IndexOf("rightend") < 0).ToDictionary(x => x.Key, y => y.Value);
            //按照行和列（先行后列）排序
            dbConfig = dbConfig.OrderBy(x =>
            Convert.ToInt32(x.Key.Substring(0, x.Key.IndexOf("-")))
            ).ThenBy(x =>
            Convert.ToInt32(x.Key.Substring(x.Key.IndexOf("-") + 1, x.Key.Length - x.Key.IndexOf("-") - 1))
            ).ToDictionary(x => x.Key, y => y.Value);
            //生成解析实例
            foreach (string dbName in DBSource.Keys)
            {
                Source source = new Source()
                {
                    sourceName = dbName,
                    oracleTableName = DBSource[dbName],
                };
                if (noInfo.ContainsKey(dbName))
                    source.NoCols = noInfo[dbName];
                if (guidInfo.ContainsKey(dbName))
                    source.GuidCols = guidInfo[dbName];
                Dictionary<string, string> dbGroup = dbConfig.Where(x => x.Value.StartsWith(dbName)|| new Regex(@"[;；]\s*" + dbName).IsMatch(x.Value)).ToDictionary(x => x.Key, y => y.Value);
                if (dbGroup.Count(x => x.Value.ToLower().IndexOf("[down]") >= 0) > 0)
                {
                    source.sourceType = SourceType.Down;
                }
                else if (dbGroup.Count(x => x.Value.ToLower().IndexOf("[right]") >= 0) > 0)
                {
                    source.sourceType = SourceType.Right;
                }
                else
                {
                    source.sourceType = SourceType.Common;
                }
                List<SourceCol> sourceList = new List<SourceCol>() { };
                foreach (var item in dbGroup)
                {
                    int pointIndex = item.Value.IndexOf(".");
                    if (pointIndex == -1 || pointIndex + 1 == item.Value.Length)
                    {
                        MessageBox.Show("配置错误，" + GetPositionByKey(item.Key) + "列名为空");
                        report("配置错误，" + GetPositionByKey(item.Key) + "列名为空");
                        return false;
                    }
                    int firstLeftBracketIndex = item.Value.IndexOf("[");
                    if (firstLeftBracketIndex == -1)
                    {
                        firstLeftBracketIndex = item.Value.Length;
                    }
                    SourceCol sourceCol = new SourceCol()
                    {
                        colName = item.Value.Substring(pointIndex + 1, firstLeftBracketIndex - pointIndex - 1),
                        position = item.Key,
                        isPublic = item.Value.ToLower().IndexOf("[public]") >= 0 ? true : false,
                        isFixed = item.Value.ToLower().IndexOf("[fixed]") >= 0 ? true : false,
                        isDate = new Regex(@"\[date\(?.*\)?\]").Matches(item.Value.ToLower()).Count > 0 ? true : false,
                    };
                    //日期格式判断
                    if (sourceCol.isDate)
                        sourceCol.dateFormat = new Regex(@"(?<=\[date\("")[^\(\)""]+(?=""\))", RegexOptions.IgnoreCase).Match(item.Value).Value;
                    if (sourceCol.isDate && string.IsNullOrWhiteSpace(sourceCol.dateFormat))
                        sourceCol.dateFormat = "yyyy/MM/dd";
                    //将第一个不是public且含有key的列名设为主键
                    if (string.IsNullOrEmpty(source.colKey) && !sourceCol.isPublic && item.Value.ToLower().IndexOf("[key]") >= 0)
                        source.colKey = sourceCol.colName;
                    //是否有空停止
                    if (item.Value.ToLower().IndexOf("[hasempty]") >= 0)
                        source.stopType = StopType.StopWhileHasEmpty;
                    //是否全空停止
                    if (item.Value.ToLower().IndexOf("[allempty]") >= 0)
                        source.stopType = StopType.StopWhileAllEmpty;
                    //是否关键词停止
                    Regex keyReg = new Regex(@"\[stopwithword=.+\]",RegexOptions.IgnoreCase);
                    if (keyReg.IsMatch(item.Value.ToLower()))
                    {
                        string matchValue = keyReg.Match(item.Value.ToLower()).Value;
                        source.stopType = StopType.StopWithWord;
                        int startIndex = matchValue.IndexOf("[stopwithword=");
                        int endIndex = matchValue.IndexOf("]");
                        source.stopKey = matchValue.Substring(startIndex+14, endIndex - startIndex-14);
                    }
                    sourceList.Add(sourceCol);
                }
                source.sourceCols = sourceList;
                //没有key的列名将第一个非public列设为主键
                if (string.IsNullOrEmpty(source.colKey) && source.sourceCols.Count(x => !x.isPublic) > 0)
                    source.colKey = source.sourceCols.First(x => !x.isPublic).colName;
                //否则将第一个列设为主键
                if (string.IsNullOrEmpty(source.colKey) && source.sourceCols.Count() > 0)
                    source.colKey = source.sourceCols.First().colName;
                result.Add(source);
            }
            //检测重复列名，列名不允许重复
            foreach (Source source in result)
            {
                var sourceColsGroupByName = source.sourceCols.GroupBy(x => x.colName);
                foreach (var item in sourceColsGroupByName)
                {
                    if (item.Count() > 1)
                    {
                        MessageBox.Show(source.sourceName + "列" + item.Key + "存在重复请检查");
                        report(source.sourceName + "列" + item.Key + "存在重复请检查");
                        return false;
                    }
                }
            }
            //检测down是否在一行，right是否在一列
            //foreach (Source source in result)
            //{
            //    if (source.sourceType == SourceType.Common)
            //    {
            //        continue;
            //    }
            //    if (source.sourceType == SourceType.Down)
            //    {
            //        var sourceColsGroupByName = source.sourceCols.GroupBy(x => x.rowIndex);
            //        if (sourceColsGroupByName.Count() > 1)
            //        {
            //            MessageBox.Show(source.sourceName + "向下自增必须处于一行");
            //            report(source.sourceName + "向下自增必须处于一行");
            //            return false;
            //        }
            //    }
            //    else if (source.sourceType == SourceType.Right)
            //    {
            //        var sourceColsGroupByName = source.sourceCols.GroupBy(x => x.colIndex);
            //        if (sourceColsGroupByName.Count() > 1)
            //        {
            //            MessageBox.Show(source.sourceName + "向右自增必须处于一列");
            //            report(source.sourceName + "向右自增必须处于一列");
            //            return false;
            //        }
            //    }
            //}
            //自增停止方式判定
            //foreach (Source source in result)
            //{
            //    //是否关键词停止
            //    if (source.sourceType == SourceType.Down)
            //    {
            //        int index = contentConfig.Values.ToList().FindIndex(x => x.ToLower().IndexOf(source.sourceName + ".downend") >= 0);
            //        int firstIndex = dbConfig.Values.ToList().FindIndex(x => x.ToLower().IndexOf(source.sourceName) >= 0);
            //        if (index == -1)
            //        {
            //            continue;
            //        }
            //        string row1Key = contentConfig.Keys.ToList()[index];
            //        string row2Key = dbConfig.Keys.ToList()[firstIndex];
            //        string row1Value = contentConfig.Values.ToList()[index];
            //        row1Value = row1Value.Replace("（", "(").Replace("）", ")");
            //        Regex regex = new Regex(@".*(?=[\(\[]" + source.sourceName + @"\.downend[\)\]])", RegexOptions.IgnoreCase);
            //        Match regMatch = regex.Match(row1Value);
            //        if (regMatch.Value.Trim() == "")
            //        {
            //            MessageBox.Show(GetPositionByKey(row1Key) + "自增结束标识前必须有关键词");
            //            report(GetPositionByKey(row1Key) + "自增结束标识前必须有关键词");
            //            return false;
            //        }
            //        source.stopKey = regMatch.Value;
            //        int rowIndex1 = Convert.ToInt32(row1Key.Substring(0, row1Key.IndexOf("-")));
            //        int rowIndex2 = Convert.ToInt32(row2Key.Substring(0, row2Key.IndexOf("-")));
            //        if (rowIndex2 < rowIndex1)
            //        {
            //            source.stopType = StopType.StopWithSymbol;
            //            continue;
            //        }
            //    }
            //    else if (source.sourceType == SourceType.Right)
            //    {
            //        int index = contentConfig.Values.ToList().FindIndex(x => x.ToLower().IndexOf(source.sourceName + ".rightend") >= 0);
            //        int firstIndex = dbConfig.Values.ToList().FindIndex(x => x.ToLower().IndexOf(source.sourceName) >= 0);
            //        if (index == -1)
            //        {
            //            continue;
            //        }
            //        string row1Key = contentConfig.Keys.ToList()[index];
            //        string row2Key = dbConfig.Keys.ToList()[firstIndex];
            //        string row1Value = contentConfig.Values.ToList()[index];
            //        row1Value = row1Value.Replace("（", "(").Replace("）", ")");
            //        Regex regex = new Regex(@".*(?=[\(\[]" + source.sourceName + @"\.rightend[\)\]])", RegexOptions.IgnoreCase);
            //        Match regMatch = regex.Match(row1Value);
            //        if (regMatch.Value.Trim() == "")
            //        {
            //            MessageBox.Show(GetPositionByKey(row1Key) + "自增结束标识前必须有关键词");
            //            report(GetPositionByKey(row1Key) + "自增结束标识前必须有关键词");
            //            return false;
            //        }
            //        source.stopKey = regMatch.Value;
            //        int rowIndex1 = Convert.ToInt32(row1Key.Substring(row1Key.IndexOf("-") + 1, row1Key.Length - row1Key.IndexOf("-") - 1));
            //        int rowIndex2 = Convert.ToInt32(row2Key.Substring(row2Key.IndexOf("-") + 1, row2Key.Length - row2Key.IndexOf("-") - 1));
            //        if (rowIndex2 < rowIndex1)
            //        {
            //            source.stopType = StopType.StopWithSymbol;
            //            continue;
            //        }
            //    }
            //}
            report("解析完成");
            return true;
        }

        private static bool VaildConfig(List<string> keys, string value)
        {
            value = value.Replace("（", "(").Replace("）", ")").Replace("【", "[").Replace("】", "]").ToLower();
            string regKeys = string.Format("({0})", string.Join("|", keys));
            Regex reg1 = new Regex(@"^" + @regKeys + @"\.[^\[\]]+(\[.+\])*$", RegexOptions.IgnoreCase);
            Regex reg2 = new Regex(@"^.*[\(\[]" + @regKeys + @"\.(downend|rightend)[\)\]]$", RegexOptions.IgnoreCase);
            if (reg1.Match(value).Success)
            {
                Regex reg = new Regex(@"(?<=\[)[^\[\]]*(?=\])");
                MatchCollection mc = reg.Matches(value);
                foreach (Match item in mc)
                {
                    if (item.Value.Trim() != "public" && item.Value.Trim() != "allempty" && item.Value.Trim() != "hasempty" && item.Value.Trim() != "down" && item.Value.Trim() != "right" && item.Value.Trim() != "key" && item.Value.Trim() != "fixed" && item.Value.Trim() != "date" && !new Regex(@"^date\(.*\)$").IsMatch(item.Value.Trim())&& !new Regex(@"^stopwithword=.+$").IsMatch(item.Value.Trim()))
                        return false;
                }
                return true;
            }
            if (reg2.Match(value).Success)
                return true;
            return false;
        }

        public static bool GetDataByAnalysis(string importFilePath, Dictionary<string, string> contentConfig, List<Source> sources, ProcessReport report, ProcessReportClear reportClear, out List<DataTable> dataTables)
        {
            dataTables = new List<DataTable>() { };
            if (contentConfig == null || contentConfig.Count == 0 || sources == null || sources.Count == 0)
            {
                MessageBox.Show("规则未解析");
                return false;
            }
            reportClear();
            report("开始读取导入文件");
            ExcelHelper excelHelper = new ExcelHelper(importFilePath);
            report("开始读取所有内容");
            //读取所有内容，合并单元格的内容都赋值第一个值
            Dictionary<string, string> allContent = excelHelper.ReadAllContent();
            if (allContent.Count == 0)
            {
                MessageBox.Show("导入文件为空");
                report("导入文件为空");
                return false;
            }
            report("开始解析规则匹配数据");
            //检测关键词结束的源向下或向后是否有关键词
            foreach (Source source in sources)
            {
                if (source.sourceType != SourceType.Common && source.stopType == StopType.StopWithWord)
                {
                    int count = 0;
                    var stopKeyList = allContent.Where(x => x.Value.ToLower().Contains(source.stopKey.ToLower())).ToList();
                    if (stopKeyList.Count == 0)
                    {
                        MessageBox.Show(source.sourceName + (source.sourceType == SourceType.Down ? "向下" : "向右") + "未找到关键词\"" + source.stopKey + "\"");
                        report(source.sourceName + (source.sourceType == SourceType.Down ? "向下" : "向右") + "未找到关键词\"" + source.stopKey + "\"");
                        return false;
                    }
                    int compIndex = -1;
                    int[] indexCollectionIn;
                    int index1;
                    int index2;
                    if (source.sourceType == SourceType.Down)
                    {
                        compIndex = source.firstRowNo;
                        indexCollectionIn = source.colIndexCollection;
                        index1 = 0;
                        index2 = 1;
                    }
                    else
                    {
                        compIndex = source.firstColNo;
                        indexCollectionIn = source.rowIndexCollection;
                        index1 = 1;
                        index2 = 0;
                    }
                    foreach (var item in stopKeyList)
                    {
                        int compIndex1 = GetPositionIndexByKey(item.Key)[index1];
                        int compIndex2 = GetPositionIndexByKey(item.Key)[index2];
                        if (compIndex1 > compIndex && indexCollectionIn.Contains(compIndex2))
                        {
                            count++;
                        }
                    }
                    if (count == 0)
                    {
                        MessageBox.Show(source.sourceName + (source.sourceType == SourceType.Down ? "向下" : "向右") + "未找到关键词\"" + source.stopKey + "\"");
                        report(source.sourceName + (source.sourceType == SourceType.Down ? "向下" : "向右") + "未找到关键词\"" + source.stopKey + "\"");
                        return false;
                    }
                }
            }
            //复制一份数据源
            List<Source> allSourceTemp = sources.ToList();
            Dictionary<Source, List<List<string>>> data = new Dictionary<Source, List<List<string>>>() { };
            //从第一个数据源开始计算修正 rowIndex 和 colIndex 并获取数据（核心代码待完善）
            bool isTimeOut = false;
            try
            {
                isTimeOut = TimeOutClass.CallWithTimeout(() =>
                {
                    FixRowAndColAndGetData(sources.First(), allSourceTemp, allContent, ref data, false, false, 0, 0);
                }, 5000);
            }
            catch (Exception e)
            {
                MessageBox.Show("提取数据失败：" + e.Message);
                report("提取数据失败：" + e.Message);
            }
            if (isTimeOut)
            {
                MessageBox.Show("提取数据失败：操作已超时，请检查文件和配置");
                report("提取数据失败：操作已超时，请检查文件和配置");
                return false;
            }
            report("导入文件数据提取完成");
            report("开始生成预览数据");
            //生成预览数据
            foreach (var item in data)
            {
                DataTable dt = new DataTable();
                dt.TableName = item.Key.oracleTableName + "(" + item.Key.sourceName + ")";
                Source source = item.Key;
                int guidNum = 0;
                if (source.GuidCols != null)
                {
                    guidNum = source.GuidCols.Count();
                    foreach (string col in source.GuidCols)
                    {
                        dt.Columns.Add(col);
                    }
                }
                int noNum = 0;
                if (source.NoCols != null)
                {
                    noNum = source.NoCols.Count();
                    foreach (string col in source.NoCols)
                    {
                        dt.Columns.Add(col);
                    }
                }
                foreach (SourceCol col in source.sourceCols)
                    dt.Columns.Add(col.colName);

                List<List<string>> dataRows = item.Value;
                for (int i = 0; i < dataRows.Count; i++)
                {
                    List<string> row = dataRows[i];
                    DataRow dr = dt.NewRow();
                    //过滤掉全部为空的数据
                    bool allRowEmpty = true;
                    for (int j = 0; j < row.Count; j++)
                    {
                        if (row[j].Trim() != string.Empty)
                        {
                            allRowEmpty = false;
                            break;
                        }
                    }
                    if (!allRowEmpty)
                    {
                        for (int j = 0; j < row.Count + guidNum + noNum; j++)
                        {
                            if (j < guidNum)
                            {
                                dr[j] = Guid.NewGuid().ToString();
                                continue;
                            }
                            if (j < guidNum + noNum)
                            {
                                dr[j] = i + 1;
                                continue;
                            }
                            int index = j - guidNum - noNum;
                            if (source.sourceCols[index].isDate)
                            {
                                try
                                {
                                    dr[j] = Convert.ToDateTime(row[index]).ToString(source.sourceCols[index].dateFormat);
                                }
                                catch (Exception ex)
                                {
                                    report("日期格式错误：" + ex.Message);
                                    dr[j] = row[index];
                                }
                            }
                            else
                            {
                                dr[j] = row[index];
                            }
                        }
                        dt.Rows.Add(dr);
                    }
                }
                dataTables.Add(dt);
            }
            report("预览数据已生成");
            return true;
        }
        private static void FixRowAndColAndGetData(Source curSource, List<Source> allSourceTemp, Dictionary<string, string> allContent, ref Dictionary<Source, List<List<string>>> dataCollection, bool downEffected, bool rightEffected, int rowInc, int colInc)
        {
            //去除当前数据源
            allSourceTemp.Remove(curSource);
            List<List<string>> data = new List<List<string>>() { };
            List<Source> effectedDownDs = new List<Source>() { };
            List<Source> effectedRightDs = new List<Source>() { };
            //获取数据
            switch (curSource.sourceType)
            {
                case SourceType.Down:
                    data = GetDownData(curSource, allContent, downEffected, rightEffected, rowInc, colInc);
                    rowInc += data.Count == 0 ? 1 : data.Count - 1;
                    effectedDownDs = GetEffectedDownDs(curSource, allSourceTemp);
                    break;
                case SourceType.Right:
                    data = GetRightData(curSource, allContent, downEffected, rightEffected, rowInc, colInc);
                    colInc += data.Count == 0 ? 1 : data.Count - 1;
                    effectedRightDs = GetEffectedRightDs(curSource, allSourceTemp);
                    break;
                case SourceType.Common:
                    data = GetCommonData(curSource, allContent, downEffected, rightEffected, rowInc, colInc);
                    break;
                default:
                    break;
            }
            //结果加入数据
            dataCollection.Add(curSource, data);
            //无数据源即结束
            if (allSourceTemp.Count == 0)
            {
                return;
            }
            downEffected = effectedDownDs.Contains(allSourceTemp.First());
            rightEffected = effectedRightDs.Contains(allSourceTemp.First());
            //下一个数据源继续解析
            FixRowAndColAndGetData(allSourceTemp.First(), allSourceTemp, allContent, ref dataCollection, downEffected, rightEffected, rowInc, colInc);
        }

        private static List<Source> GetEffectedRightDs(Source source, List<Source> allSource)
        {
            List<Source> listSourceName = new List<Source>() { };
            //满足通用模板需完善该处代码
            foreach (Source item in allSource)
            {
                if (item.colIndexFullCollection.Count(x => x >= source.firstColNo) > 0)
                {
                    listSourceName.Add(item);
                }
            }
            return listSourceName;
        }

        private static List<Source> GetEffectedDownDs(Source source, List<Source> allSource)
        {
            List<Source> listSourceName = new List<Source>() { };
            //满足通用模板需完善该处代码
            foreach (Source item in allSource)
            {
                if (item.rowIndexFullCollection.Count(x => x >= source.firstRowNo) > 0)
                {
                    listSourceName.Add(item);
                }
            }
            return listSourceName;
        }
        private static List<List<string>> GetCommonData(Source source, Dictionary<string, string> allContent, bool downEffected, bool rightEffected, int rowInc, int colInc)
        {
            List<List<string>> result = new List<List<string>>() { };
            List<string> values = GetCommonValues(allContent, source, downEffected, rightEffected, rowInc, colInc);
            result.Add(values);
            return result;
        }

        private static List<List<string>> GetDownData(Source source, Dictionary<string, string> allContent, bool downEffected, bool rightEffected, int rowInc, int colInc)
        {
            List<List<string>> result = new List<List<string>>() { };
            if (source.stopType == StopType.StopWithWord)
            {
                int rowIndex = source.sourceCols.Find(x => x.colName == source.colKey).rowIndex;
                do
                {
                    List<string> values = GetDownValues(allContent, rowIndex, source, downEffected, rightEffected, rowInc, colInc);
                    rowIndex++;
                    //关键词结束
                    if (rowIndex > source.firstRowNo + 1 && (values.Count == 0 || values.Count(x => x.IndexOf(source.stopKey) >= 0) > 0))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            else if (source.stopType == StopType.StopWhileHasEmpty)
            {
                int rowIndex = source.sourceCols.Find(x => x.colName == source.colKey).rowIndex;
                do
                {
                    List<string> values = GetDownValues(allContent, rowIndex, source, downEffected, rightEffected, rowInc, colInc);
                    rowIndex++;
                    //读到有空结束
                    if (rowIndex > source.firstRowNo + 1 && (values.Count == 0 || values.Count(x => x.Trim() == "") > 0))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            else if (source.stopType == StopType.StopWhileAllEmpty)
            {
                int rowIndex = source.sourceCols.Find(x => x.colName == source.colKey).rowIndex;
                do
                {
                    List<string> values = GetDownValues(allContent, rowIndex, source, downEffected, rightEffected, rowInc, colInc);
                    rowIndex++;
                    List<string> checkValues = new List<string>();
                    for (int i = 0; i < source.sourceCols.Count; i++)
                    {
                        var s =  source.sourceCols[i];
                        if (!s.isFixed&&!s.isPublic)
                        {
                            checkValues.Add(values[i]);
                        }
                    }
                    //读到全空结束
                    if (rowIndex > source.firstRowNo + 1 && (checkValues.Count == 0 || checkValues.Count(x => x.Trim() != "") == 0))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            else if (source.stopType == StopType.StopWhileKeyEmpty)
            {
                int rowIndex = source.sourceCols.Find(x => x.colName == source.colKey).rowIndex;
                do
                {
                    List<string> values = GetDownValues(allContent, rowIndex, source, downEffected, rightEffected, rowInc, colInc);
                    rowIndex++;
                    //读到主键为空结束
                    if (rowIndex > source.firstRowNo + 1 && (values.Count == 0 || values[source.sourceCols.FindIndex(x => x.colName == source.colKey)] == ""))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            return result;
        }
        private static List<List<string>> GetRightData(Source source, Dictionary<string, string> allContent, bool downEffected, bool rightEffected, int rowInc, int colInc)
        {
            List<List<string>> result = new List<List<string>>() { };
            if (source.stopType == StopType.StopWithWord)
            {
                int colIndex = source.firstNotFixedColNo;
                do
                {
                    List<string> values = GetRightValues(allContent, colIndex, source, downEffected, rightEffected, rowInc, colInc);
                    colIndex++;
                    //关键词结束
                    if (colIndex > source.firstColNo + 1 && (values.Count == 0 || values.Count(x => x.ToLower().IndexOf(source.stopKey) >= 0) > 0))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            else if (source.stopType == StopType.StopWhileHasEmpty)
            {
                int colIndex = source.firstNotFixedColNo;
                do
                {
                    List<string> values = GetRightValues(allContent, colIndex, source, downEffected, rightEffected, rowInc, colInc);
                    colIndex++;
                    //读到有空结束
                    if (colIndex > source.firstColNo + 1 && (values.Count == 0 || values.Count(x => x.Trim() == "") > 0))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            else if (source.stopType == StopType.StopWhileAllEmpty)
            {
                int colIndex = source.firstNotFixedColNo;
                do
                {
                    List<string> values = GetRightValues(allContent, colIndex, source, downEffected, rightEffected, rowInc, colInc);
                    colIndex++;
                    List<string> checkValues = new List<string>();
                    for (int i = 0; i < source.sourceCols.Count; i++)
                    {
                        var s = source.sourceCols[i];
                        if (!s.isFixed && !s.isPublic)
                        {
                            checkValues.Add(values[i]);
                        }
                    }
                    //读到全空结束
                    if (colIndex > source.firstColNo + 1 && (checkValues.Count == 0 || checkValues.Count(x => x.Trim() != "") == 0))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            else if (source.stopType == StopType.StopWhileKeyEmpty)
            {
                int colIndex = source.firstNotFixedColNo;
                do
                {
                    List<string> values = GetRightValues(allContent, colIndex, source, downEffected, rightEffected, rowInc, colInc);
                    colIndex++;
                    //读到主键为空结束
                    if (colIndex > source.firstColNo + 1 && (values.Count == 0 || values[source.sourceCols.FindIndex(x => x.colName == source.colKey)] == ""))
                    {
                        break;
                    }
                    result.Add(values);
                } while (true);
            }
            return result;
        }

        private static List<string> GetCommonValues(Dictionary<string, string> allContent, Source source, bool downEffected, bool rightEffected, int rowInc, int colInc)
        {
            List<string> result = new List<string>() { };
            for (int i = 0; i < source.sourceCols.Count; i++)
            {
                SourceCol item = source.sourceCols[i];
                int rowIndex = item.rowIndex;
                int colIndex = item.colIndex;
                //满足通用模板需完善该处代码
                rowIndex += rowInc;
                colIndex += colInc;
                string key = string.Format("{0}-{1}", rowIndex, colIndex);
                if (item.isFixed)
                    key = string.Format("{0}-{1}", item.rowIndex, item.colIndex);
                if (allContent.ContainsKey(key))
                {
                    result.Add(allContent[key]);
                }
            }
            return result;
        }

        private static List<string> GetRightValues(Dictionary<string, string> allContent, int colIndex, Source source, bool downEffected, bool rightEffected, int rowInc, int colInc)
        {
            List<string> result = new List<string>() { };
            colIndex += colInc;
            for (int i = 0; i < source.sourceCols.Count; i++)
            {
                SourceCol item = source.sourceCols[i];
                int rowIndex = item.rowIndex;
                //满足通用模板需完善该处代码
                rowIndex += rowInc;
                string key = string.Format("{0}-{1}", rowIndex, colIndex);
                if (item.isPublic)
                    key = string.Format("{0}-{1}", rowIndex, source.firstColNo + colInc);
                if (item.isFixed)
                    key = string.Format("{0}-{1}", item.rowIndex, item.colIndex);
                if (allContent.ContainsKey(key))
                {
                    result.Add(allContent[key]);
                }
                else
                {
                    result.Add("");
                }
            }
            return result;
        }

        private static List<string> GetDownValues(Dictionary<string, string> allContent, int rowIndex, Source source, bool downEffected, bool rightEffected, int rowInc, int colInc)
        {
            List<string> result = new List<string>() { };
            rowIndex += rowInc;
            for (int i = 0; i < source.sourceCols.Count; i++)
            {
                SourceCol item = source.sourceCols[i];
                int colIndex = item.colIndex;

                //满足通用模板需完善该处代码
                colIndex += colInc;
                string key = string.Format("{0}-{1}", rowIndex, colIndex);
                if (item.isPublic)
                    key = string.Format("{0}-{1}", source.firstRowNo + rowInc, colIndex);
                if (item.isFixed)
                    key = string.Format("{0}-{1}", source.firstRowNo, item.colIndex);
                if (allContent.ContainsKey(key))
                {
                    result.Add(allContent[key]);
                }
                else
                {
                    result.Add("");
                }
            }
            return result;
        }

        public static string GetPositionByKey(string key)
        {
            int index = key.IndexOf("-");
            return ExcelHelper.NumToAlpha(Convert.ToInt32(key.Substring(index + 1, key.Length - index - 1)), false).ToString() + (Convert.ToInt32(key.Substring(0, index))).ToString();
        }

        public static int[] GetPositionIndexByKey(string key)
        {
            int[] result = new int[2];
            int index = key.IndexOf("-");
            result[0] = Convert.ToInt32(key.Substring(0, index));
            result[1] = Convert.ToInt32(key.Substring(index + 1, key.Length - index - 1));
            return result;
        }
    }
}
