using ExcelHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml;
using System.Collections;
using System.Reflection;

namespace Excel2Oracle
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public static string templatesDirectoryName = "Templates";
        public static string exeDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

        public static string templateFilePath = "";
        public static string importedFilePath = "";
        public static AnalysisHelper.ProcessReport report;
        public static AnalysisHelper.InsertProcessReport insertReport;
        public static AnalysisHelper.ProcessReportClear reportClear;
        public static AnalysisHelper.InsertProcessReportClear insertReportClear;

        List<Source> result;
        Dictionary<string, string> contentConfig;
        List<DataTable> dataTables;
        Dictionary<string, string[]> noInfo;
        Dictionary<string, string[]> guidInfo;
        Dictionary<string, string> beforeInfo;
        Dictionary<string, string> afterInfo;
        OrclConfig orclConfig = new OrclConfig();
        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainLoad;
            this.WindowState = WindowState.Maximized;
            //txtFilePath.Text = System.IO.Path.Combine(exeDirectory, @"D:\Documents\Visual Studio 2015\Projects\Excel2Oracle\07 Excel导入Oracle工具\Excel2Oracle\bin\Debug\导入文件\201809-阳高蔚家堡环境监测报表.xls");
            //txtFilePath.Text = @"D:\Documents\Visual Studio 2015\Projects\Excel2Oracle\07 Excel导入Oracle工具\Excel2Oracle\bin\Debug\导入文件\201809-阳高蔚家堡环境监测报表.xls";
        }


        private void MainLoad(object sender, RoutedEventArgs e)
        {
            this.Title = "报表数据-Oracle导入工具";
            //写进度委托加入事件
            report += WriteLineProcess;
            insertReport += WriteLineInsertProcess;
            //清理进度委托加入事件
            reportClear += ClearProcess;
            insertReportClear += ClearInsertProcess;
            //加载模板树
            TemplatesTreeViewLoad();
        }

        private void TemplatesTreeViewLoad()
        {
            string templatesPath = System.IO.Path.Combine(exeDirectory, templatesDirectoryName);
            if (!System.IO.Directory.Exists(templatesPath))
            {
                System.Windows.MessageBox.Show("未找到模板目录");
                return;
            }
            TemplateTreeViewItem[] allTreeNodes = ExcelFilesHelper.GetAllExcelFiles(templatesPath);
            TemplateTreeViewItem root = new TemplateTreeViewItem(NodeType.Root, templatesPath, allTreeNodes, "模板");
            root.Icon = "Images/templateRoot.jpg";
            root.IsExpanded = true;
            root.Visiblity = "Visible";
            TemplateTreeView.ItemsSource = new TemplateTreeViewItem[] { root };
        }

        private void TemplateTreeView_SelectedItemChanged(object sender, EventArgs e)
        {
            TemplateTreeViewItem item = (TemplateTreeViewItem)TemplateTreeView.SelectedItem;
            if (item.NodeType == NodeType.File)
            {
                lblCurTemplate.Content = item.DisplayName;
                templateFilePath = item.Path;
                tipCurTemplate.Text = item.DisplayName;
                AnalysisFile(null,null);
            }
            else
            {
                lblCurTemplate.Content = "未选择";
                templateFilePath = "";
                tipCurTemplate.Text = "未选择";
            }
        }

        private void ReloadTemplates(object sender, RoutedEventArgs e)
        {
            TemplatesTreeViewLoad();
        }

        private void SelectImportedFile(object sender, EventArgs e)
        {
            Microsoft.Win32.OpenFileDialog ofd = new Microsoft.Win32.OpenFileDialog();
            ofd.DefaultExt = ".xls";
            ofd.Filter =
              "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            if (ofd.ShowDialog() == true)
            {
                txtFilePath.Text = ofd.FileName;
            }
        }
        /// <summary>
        /// 解析规则
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AnalysisFile(object sender, EventArgs e)
        {
            if (templateFilePath == "")
            {
                MessageBox.Show("模板文件未选择");
                return;
            }
            if (!File.Exists(templateFilePath))
            {
                MessageBox.Show("模板文件不存在");
                return;
            }
            //解析规则
            bool success = AnalysisHelper.Analysis(templateFilePath, report, reportClear, out result, out contentConfig, out guidInfo, out noInfo, out beforeInfo, out afterInfo);
            if (success)
            {
                lblCurRule.Content = lblCurTemplate.Content;
                tipCurRule.Text = lblCurTemplate.Content as string;
            }

        }
        public void WriteLineProcess(string info)
        {

            txtProcess.Text = string.Format("{0}{1}{2}", txtProcess.Text, txtProcess.Text.Trim() == "" ? "" : "\r\n", info);
            txtProcess.ScrollToEnd();
        }
        public void WriteLineInsertProcess(string info)
        {
            txtImportProcess.Text = string.Format("{0}{1}{2}", txtImportProcess.Text, txtImportProcess.Text.Trim() == "" ? "" : "\r\n", info);
            txtImportProcess.ScrollToEnd();
        }
        public void ClearProcess()
        {
            txtProcess.Text = "";
        }
        public void ClearInsertProcess()
        {
            txtImportProcess.Text = "";
        }

        /// <summary>
        /// 导入数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Import(object sender, EventArgs e)
        {
            insertReportClear();
            if (dataTables == null)
            {
                insertReport("请先提取数据");
                MessageBox.Show("请先提取数据");
                return;
            }
            if (dataTables.Count == 0 || dataTables.Count(x => x.Rows.Count > 0) == 0)
            {
                insertReport("未找到任何数据");
                MessageBox.Show("未找到任何数据");
                return;
            }
            try
            {
                for (int k = 0; k < dataTables.Count; k++)
                {
                    DataTable table = dataTables[k];
                    insertReport("正在检测表" + table.TableName);
                    if (table.Rows.Count == 0)
                    {
                        insertReport("表" + table + "不存在数据");
                        continue;
                    }
                    //获取表名
                    string subname = result[k].sourceName;
                    string tableName = result[k].oracleTableName;
                    //建列名
                    string createColumns = "";
                    string insertColumns = "";
                    Dictionary<string, int> colMaxLens = new Dictionary<string, int>() { };
                    List<string> willAddColumns = new List<string>();
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        DataColumn col = table.Columns[i];
                        List<int> colMaxLen = new List<int>() { };
                        for (int j = 0; j < table.Rows.Count; j++)
                        {
                            colMaxLen.Add(table.Rows[j][i].ToString().Length);
                        }
                        int maxLen = colMaxLen.Max();
                        colMaxLens.Add(col.ColumnName, (maxLen / 50) * 50 + 50);
                    }
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        DataColumn col = table.Columns[i];
                        var obj = result[k].sourceCols.Find(x => x.colName == table.Columns[i].ColumnName);
                        if (obj != null && obj.isDate)
                        {
                            createColumns += "\"" + col.ColumnName + "\"" + " date,";
                        }
                        else
                        {
                            createColumns += "\"" + col.ColumnName + "\"" + " varchar(" + colMaxLens[col.ColumnName] + "),";
                        }
                        insertColumns += "\"" + col.ColumnName + "\",";
                    }
                    createColumns = createColumns.Remove(createColumns.Length - 1, 1);
                    insertColumns = insertColumns.Remove(insertColumns.Length - 1, 1);
                    //判断是否存在表名
                    string existTableName = "";
                    if (!orclConfig.EnabledInterface)
                    {
                        existTableName = OracleHelper.ExecuteScalar(orclConfig.dBConnectInfo, CommandType.Text, "select table_name from user_tables where table_name='" + tableName + "'").ToString();
                    }
                    else
                    {
                        object[][] data = orclConfig.sqlClient.DoSelectMethod("select table_name from user_tables where table_name='" + tableName + "'");
                        if (data[0] != null && data[0][0] != null)
                        {
                            existTableName = data[0][0] as string;
                        }
                    }

                    string insertSqls = "begin ";//单张表的总的插入语句
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        string insertSql = "insert into {0}({1}) values({2});";
                        DataRow row = table.Rows[i];

                        StringBuilder insertValues = new StringBuilder();
                        for (int j = 0; j < row.ItemArray.Count(); j++)
                        {
                            var item = row.ItemArray[j];
                            var obj = result[k].sourceCols.Find(x => x.colName == table.Columns[j].ColumnName);
                            if (obj != null && obj.isDate)
                            {
                                insertValues.AppendFormat("to_date('{0}','{1}'),", item, obj.dateFormat.Replace("HH", "hh24").Replace("mm", "mi"));
                            }
                            else
                            {
                                insertValues.AppendFormat("'{0}',", item);
                            }
                        }
                        insertValues.Remove(insertValues.Length - 1, 1);
                        insertSql = string.Format(insertSql, tableName, insertColumns, insertValues);
                        insertSqls += insertSql;
                    }
                    insertSqls += "end;";
                    if (existTableName != "") //存在该表
                    {
                        if (!ExecuteBefore(k, subname, tableName, table))
                        {
                            insertReport("取消导入...");
                            MessageBox.Show("取消导入...");
                            return;
                        }
                        insertReport("开始导入...");
                        int successNum = InsertValues(tableName, insertSqls);
                        if (successNum > 0)
                        {
                            insertReport(table.TableName + "全部导入成功");
                        }
                        else
                        {
                            MessageBox.Show(table.TableName + "导入失败");
                            insertReport(table.TableName + "导入失败");
                            return;
                        }
                        if (!ExecuteAfter(k, subname, tableName, table))
                        {
                            insertReport("取消导入...");
                            MessageBox.Show("取消导入...");
                            return;
                        }
                    }
                    else
                    {
                        //不存在该表 ，需要创建
                        string createTable = "create table " + tableName + string.Format("({0})", createColumns);
                        insertReport(createTable);
                        MessageBoxResult dr = MessageBox.Show("检测到表" + tableName + "不存在，是否创建该表？", "确认继续", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (dr != MessageBoxResult.Yes)
                        {
                            insertReport("取消导入...");
                            MessageBox.Show("取消导入...");
                            return;
                        }
                        try
                        {
                            if (!orclConfig.EnabledInterface)
                            {
                                OracleHelper.ExecuteNonQuery(orclConfig.dBConnectInfo, CommandType.Text, createTable);
                            }
                            else
                            {
                                orclConfig.sqlClient.OperateSql(createTable);
                            }
                            insertReport("创建表" + tableName + "完毕！");
                        }
                        catch (Exception ex)
                        {
                            insertReport("建表sql执行失败：\r\n" + ex.Message);
                            MessageBox.Show("建表sql执行失败：\r\n" + ex.Message);
                            return;
                        }
                        insertReport("开始导入...");
                        if (!ExecuteBefore(k, subname, tableName, table))
                        {
                            insertReport("取消导入...");
                            MessageBox.Show("取消导入...");
                            return;
                        }
                        int successNum = InsertValues(tableName, insertSqls);
                        if (successNum > 0)
                        {
                            insertReport(table.TableName + "全部导入成功");
                        }
                        else
                        {
                            MessageBox.Show(table.TableName + "导入失败");
                            insertReport(table.TableName + "导入失败");
                            return;
                        }
                        if (!ExecuteAfter(k, subname, tableName, table))
                        {
                            insertReport("取消导入...");
                            MessageBox.Show("取消导入...");
                            return;
                        }
                    }
                }
                insertReport("全部导入结束");
                MessageBox.Show("全部导入结束");
            }
            catch (Exception ex)
            {
                insertReport("导入失败：" + ex.ToString());
                MessageBox.Show("导入失败：" + ex.ToString());
                return;
            }
        }

        private bool ExecuteAfter(int k, string subname, string tableName, DataTable table)
        {
            insertReport("检测导入后sql...");
            if (afterInfo.ContainsKey(subname))
            {
                insertReport("执行导入后sql...");
                string sql = afterInfo[subname];
                sql = sql.Replace(subname, tableName);
                List<string> listCols = new List<string>() { };
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    listCols.Add(table.Columns[i].ColumnName);
                }
                string[] cols = listCols.ToArray();
                List<object> values = new List<object>() { };
                for (int i = 0; i < cols.Length; i++)
                {
                    values.Add("");
                }
                object[] vals = new object[] { };
                if (table.Rows.Count > 0)
                    vals = table.Rows[0].ItemArray.ToArray();
                for (int n = 0; n < cols.Length; n++)
                {
                    sql = sql.Replace(tableName + "." + cols[n], vals[n].ToString());
                }
                sql = sql.Trim();
                insertReport("导入后sql为：" + sql);
                MessageBoxResult dr = MessageBox.Show("导入后sql为：" + sql + "  是否执行？", "确认继续", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (dr == MessageBoxResult.Yes)
                {
                    try
                    {
                        int res = 0;
                        if (!orclConfig.EnabledInterface)
                        {
                            res =  OracleHelper.ExecuteNonQuery(orclConfig.dBConnectInfo, CommandType.Text, sql);
                        }
                        else
                        {
                            res = orclConfig.sqlClient.OperateSql(sql);
                        }
                        if (res > 0)
                        {
                            insertReport("导入后sql执行成功");
                        }
                        else
                        {
                            insertReport("导入后sql执行失败");
                        }
                    }
                    catch (Exception ex)
                    {
                        insertReport("导入后sql执行失败：\r\n" + ex.Message);
                        MessageBox.Show("导入后sql执行失败：\r\n" + ex.Message);
                        MessageBoxResult dr2 = MessageBox.Show("是否继续导入？", "确认继续", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (dr2 == MessageBoxResult.No)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    MessageBoxResult dr2 = MessageBox.Show("是否继续导入？", "确认继续", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (dr2 == MessageBoxResult.No)
                    {
                        return false;
                    }
                }
            }
            else
            {
                insertReport("未配置导入后sql...");
            }
            return true;
        }

        private bool ExecuteBefore(int k, string subname, string tableName, DataTable table)
        {
            insertReport("检测导入前sql...");
            if (beforeInfo.ContainsKey(subname))
            {
                insertReport("执行导入前sql...");
                string sql = beforeInfo[subname];
                sql = sql.Replace(subname, tableName);
                List<string> listCols = new List<string>() { };
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    listCols.Add(table.Columns[i].ColumnName);
                }
                string[] cols = listCols.ToArray();
                List<object> values = new List<object>() { };
                for (int i = 0; i < cols.Length; i++)
                {
                    values.Add("");
                }
                object[] vals = new object[] { };
                if (table.Rows.Count > 0)
                    vals = table.Rows[0].ItemArray.ToArray();
                for (int n = 0; n < cols.Length; n++)
                {
                    sql = sql.Replace(tableName + "." + cols[n], vals[n].ToString());
                }
                sql = sql.Trim();
                insertReport("导入前sql为：" + sql);
                MessageBoxResult dr = MessageBox.Show("导入前sql为：" + sql + "  是否执行？", "确认继续", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (dr == MessageBoxResult.Yes)
                {
                    try
                    {
                        int res = 0;
                        if (!orclConfig.EnabledInterface)
                        {
                            res = OracleHelper.ExecuteNonQuery(orclConfig.dBConnectInfo, CommandType.Text, sql);
                        }
                        else
                        {
                            res = orclConfig.sqlClient.OperateSql(sql);
                        }
                        if (res > 0)
                        {
                            insertReport("导入前sql执行成功");
                        }
                        else
                        {
                            insertReport("导入前sql执行失败");
                        }
                    }
                    catch (Exception ex)
                    {
                        insertReport("导入前sql执行失败：\r\n" + ex.Message);
                        MessageBox.Show("导入前sql执行失败：\r\n" + ex.Message);
                        MessageBoxResult dr2 = MessageBox.Show("是否继续导入？", "确认继续", MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (dr2 == MessageBoxResult.No)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    MessageBoxResult dr2 = MessageBox.Show("是否继续导入？", "确认继续", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    if (dr2 == MessageBoxResult.No)
                    {
                        return false;
                    }
                }
            }
            else
            {
                insertReport("未配置导入前sql...");
            }
            return true;
        }

        private int InsertValues(string tableName, string insertSqls)
        {
            insertReport("开始导入 sql:" + insertSqls);
            int result = 0;
            if (!orclConfig.EnabledInterface)
            {
                result = OracleHelper.ExecuteNonQuery(orclConfig.dBConnectInfo, CommandType.Text, insertSqls);
            }
            else
            {
                result = orclConfig.sqlClient.OperateSql(insertSqls);
            }
            if (result > 0)
            {
                insertReport("表" + tableName + result + "条数据导入成功");
            }
            return result;
        }
        /// <summary>
        /// 提取数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GetData(object sender, EventArgs e)
        {
            string importFilePath = txtFilePath.Text.Trim();
            if (importFilePath == "")
            {
                MessageBox.Show("导入文件未选择");
                return;
            }
            if (!File.Exists(importFilePath))
            {
                MessageBox.Show("导入文件不存在");
                return;
            }
            AnalysisHelper.GetDataByAnalysis(importFilePath, contentConfig, result, report, reportClear, out dataTables);

            txtViewTip.Visibility = Visibility.Collapsed;
            spView.Children.Clear();
            foreach (DataTable table in dataTables)
            {
                Label lblTableName = new Label();
                lblTableName.Foreground = Brushes.Red;
                lblTableName.Content = table.TableName;
                ListView lv = new ListView();
                lv.HorizontalAlignment = HorizontalAlignment.Stretch;
                lv.VerticalAlignment = VerticalAlignment.Stretch;
                GridView gv = new GridView();
                foreach (DataColumn col in table.Columns)
                {
                    gv.Columns.Add(new GridViewColumn() { Header = col.ColumnName, DisplayMemberBinding = new Binding() { Path = new PropertyPath(col.ColumnName) } });
                }
                lv.View = gv;
                lv.ItemsSource = table.DefaultView;
                spView.Children.Add(lblTableName);
                spView.Children.Add(lv);
            }
            viewScroll.Content = spView;
        }

        /// <summary>
        /// 查看当前规则
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ViewRules(object sender, RoutedEventArgs e)
        {
            RuleView ruleView = new RuleView(result);
            ruleView.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ruleView.ShowDialog();
        }

        private void txtFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            txtFilePath.ToolTip = txtFilePath.Text;
            importedFilePath = txtFilePath.Text;
        }
        /// <summary>
        /// 右击打开模板文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openTemplateFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFile(templateFilePath);
        }
        /// <summary>
        /// 右击打开导入文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openImportFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFile(importedFilePath);
        }
        /// <summary>
        /// 双击打开模板文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lblCurTemplate_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenFile(templateFilePath);
        }
        /// <summary>
        /// 双击打开导入文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtFilePath_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenFile(importedFilePath);
        }
        /// <summary>
        /// 右击打开模板文件目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openTemplatePath_Click(object sender, RoutedEventArgs e)
        {
            OpenFilePath(templateFilePath);
        }
        /// <summary>
        /// 右击打开导入文件目录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void openImportPath_Click(object sender, RoutedEventArgs e)
        {
            OpenFilePath(importedFilePath);
        }
        private static void OpenFile(string filePath)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                if (!File.Exists(templateFilePath))
                {
                    MessageBox.Show("文件不存在，打开失败");
                    return;
                }
                System.Diagnostics.Process.Start(filePath);
            }
        }
        private static void OpenFilePath(string filePath)
        {
            if (filePath == "")
                return;
            if (!string.IsNullOrEmpty(filePath))
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("文件目录不存在，打开失败");
                    return;
                }
                System.Diagnostics.Process.Start("Explorer.exe", @"/select," + filePath);
            }
        }

        private void OrclConfig_Click(object sender, RoutedEventArgs e)
        {
            orclConfig.ShowDialog();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }
    }
}