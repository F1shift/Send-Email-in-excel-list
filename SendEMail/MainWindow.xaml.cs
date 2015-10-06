using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace SendEMail
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Settings.Default.Save();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            OpenFileDialog.Filter = "Excel 活頁簿|*.xlsx";
            if (OpenFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.tb_ExcelFilePath.Text = OpenFileDialog.FileName;
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (this.tb_ExcelFilePath.Text != null)
            {
                Properties.Settings.Default.Save();

                using (SmtpClient client = new SmtpClient(
                    Properties.Settings.Default.smtpServer,
                    Properties.Settings.Default.smtpPort))//或公司、客戶的smtp_server
                {
                    if (!string.IsNullOrEmpty(Properties.Settings.Default.mailAccount) && !string.IsNullOrEmpty(tb_Pwd.Password))//.config有帳密的話
                    {//寄信要不要帳密？眾說紛紜Orz，分享一下經驗談....

                        //網友阿尼尼:http://www.dotblogs.com.tw/kkc123/archive/2012/06/26/73076.aspx
                        //※公司內部不用認證,寄到外部信箱要特別認證 Account & Password

                        //自家公司MIS:
                        //※要看smtp server的設定呀~

                        //結論...
                        //※程式在客戶那邊執行的話，問客戶，程式在自家公司執行的話，問自家公司MIS，最準確XD
                        client.Credentials = new NetworkCredential(Properties.Settings.Default.mailAccount, tb_Pwd.Password);//寄信帳密
                    }

                    bool succeed = MailSender.SendMails(this.tb_ExcelFilePath.Text, client);
                    if (succeed)
                    {
                        MessageBox.Show("Sending Done!");
                    }
                }//end using 
            }
        }

        
    }
}
