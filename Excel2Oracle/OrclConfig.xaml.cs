using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml;

namespace Excel2Oracle
{
    /// <summary>
    /// OrclConfig.xaml 的交互逻辑
    /// </summary>
    public partial class OrclConfig : Window
    {
        public static string exeDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        public static string connectInfoFileName = "ConnectInfo.xml";
        public string dBConnectInfo = "User Id={0};Password={1};Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST={2})(PORT={3})))(CONNECT_DATA=(SERVICE_NAME={4})))";
        public string tempDBConnectInfo = "User Id={0};Password={1};Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST={2})(PORT={3})))(CONNECT_DATA=(SERVICE_NAME={4})))";
        public BGScreen.BusinessServiceClient sqlClient;
        public bool EnabledInterface = false;
        public OrclConfig()
        {
            InitializeComponent();
            DBConnectStringLoad();
            this.Loaded += MainLoad;
        }

        private void MainLoad(object sender, RoutedEventArgs e)
        {
            DBConnectStringLoad();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        /// <summary>
        /// 保存关系库连接字符串
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SaveConnection(object sender, EventArgs e)
        {
            Regex ipAndPortReg = new Regex(@"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}:\d{1,5}$");
            if (chkEnableInterface.IsChecked.Value && !ipAndPortReg.IsMatch(txtWCFAddr.Text.Trim()))
            {
                MessageBox.Show("WCFIP格式不正确！");
                return;
            }
            string connectInfoPath = System.IO.Path.Combine(exeDirectory, connectInfoFileName);
            File.WriteAllText(connectInfoPath, string.Format("<?xml version=\"1.0\" encoding=\"utf-8\" ?><Config  UserId = \"{0}\" Password = \"{1}\" HOST = \"{2}\" PORT = \"{3}\" SERVICENAME = \"{4}\" EnableInterface = \"{5}\"  WCFAddr = \"{6}\"/>", txtUserID.Text, txtPassword.Text, txtHost.Text, txtPort.Text, txtServiceName.Text, chkEnableInterface.IsChecked, txtWCFAddr.Text));

            EnabledInterface = chkEnableInterface.IsChecked.Value;

            if (File.Exists(connectInfoPath))
            {
                MessageBox.Show("保存成功！");
            }
            if (!string.IsNullOrWhiteSpace(txtWCFAddr.Text))
            {
                sqlClient = new BGScreen.BusinessServiceClient("BasicHttpBinding_IBusinessService", "http://" + txtWCFAddr.Text.Trim() + "/BigScreen/BusinessService");
            }
            //this.Hide();
        }

        private void DBConnectStringLoad()
        {
            try
            {
                string connectInfoPath = System.IO.Path.Combine(exeDirectory, connectInfoFileName);
                if (File.Exists(connectInfoPath))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(connectInfoPath);
                    XmlAttributeCollection attrs = doc.SelectSingleNode("Config").Attributes;
                    txtUserID.Text = attrs["UserId"].Value;
                    txtPassword.Text = attrs["Password"].Value;
                    txtHost.Text = attrs["HOST"].Value;
                    txtPort.Text = attrs["PORT"].Value;
                    txtServiceName.Text = attrs["SERVICENAME"].Value;
                    string WCFIP = attrs["WCFAddr"].Value;
                    txtWCFAddr.Text = WCFIP;
                    chkEnableInterface.IsChecked = bool.Parse(attrs["EnableInterface"].Value);
                    EnabledInterface = bool.Parse(attrs["EnableInterface"].Value);
                    if (!string.IsNullOrWhiteSpace(WCFIP))
                    {
                        sqlClient = new BGScreen.BusinessServiceClient("BasicHttpBinding_IBusinessService", "http://" + WCFIP + "/BigScreen/BusinessService");
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("配置加载失败：" + e.Message);
            }
        }

        /// <summary>
        /// 测试关系库连接字符串
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TestDBConnect(object sender, EventArgs e)
        {
            if (chkEnableInterface.IsChecked.Value)
            {
                Regex ipAndPortReg = new Regex(@"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}:\d{1,5}$");
                if (!ipAndPortReg.IsMatch(txtWCFAddr.Text.Trim()))
                {
                    MessageBox.Show("WCFIP格式不正确！");
                    return;
                }
                sqlClient = new BGScreen.BusinessServiceClient("BasicHttpBinding_IBusinessService", "http://" + txtWCFAddr.Text.Trim() + "/BigScreen/BusinessService");
                if (sqlClient == null)
                {
                    MessageBox.Show("接口初始化失败，请检查WCFIP和端口是否正确");
                    return;
                }
                try
                {
                    int result = 0;
                    object[][] data = sqlClient.DoSelectMethod("select 1 from dual");
                    if (data[0] != null && data[0][0] != null)
                    {
                        result  = Convert.ToInt16(data[0][0]);
                    }
                    if (result == 1)
                    {
                        MessageBox.Show("接口连接成功！");
                    }
                    else
                    {
                        MessageBox.Show("接口连接失败");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("接口连接失败：" + ex.ToString());
                }
            }
            else
            {
                try
                {
                    string tempConnectInfo = tempDBConnectInfo.Replace("{0}", txtUserID.Text).Replace("{1}", txtPassword.Text).Replace("{2}", txtHost.Text).Replace("{3}", txtPort.Text).Replace("{4}", txtServiceName.Text);
                    bool result = 1 == OracleHelper.ConnectionTest(tempConnectInfo);
                    if (result)
                    {
                        MessageBox.Show("oracle连接成功！");
                    }
                    else
                    {
                        MessageBox.Show("oracle连接失败！");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("oracle连接失败：" + ex.ToString());
                }
            }
        }

        private void txtUserID_TextChanged(object sender, TextChangedEventArgs e)
        {
            dBConnectInfo = dBConnectInfo.Replace("{0}", txtUserID.Text);
        }

        private void txtPassword_TextChanged(object sender, TextChangedEventArgs e)
        {
            dBConnectInfo = dBConnectInfo.Replace("{1}", txtPassword.Text);
        }

        private void txtHost_TextChanged(object sender, TextChangedEventArgs e)
        {
            dBConnectInfo = dBConnectInfo.Replace("{2}", txtHost.Text);
        }

        private void txtPort_TextChanged(object sender, TextChangedEventArgs e)
        {
            dBConnectInfo = dBConnectInfo.Replace("{3}", txtPort.Text);
        }

        private void txtServiceName_TextChanged(object sender, TextChangedEventArgs e)
        {
            dBConnectInfo = dBConnectInfo.Replace("{4}", txtServiceName.Text);
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void chkEnableInterface_Checked(object sender, RoutedEventArgs e)
        {
            txtHost.IsEnabled = false;
            txtPassword.IsEnabled = false;
            txtServiceName.IsEnabled = false;
            txtUserID.IsEnabled = false;
            txtPort.IsEnabled = false;
            txtWCFAddr.IsEnabled = true;
        }

        private void chkEnableInterface_Unchecked(object sender, RoutedEventArgs e)
        {
            txtHost.IsEnabled = true;
            txtPassword.IsEnabled = true;
            txtServiceName.IsEnabled = true;
            txtUserID.IsEnabled = true;
            txtPort.IsEnabled = true;
            txtWCFAddr.IsEnabled = false;
        }

    }
}
