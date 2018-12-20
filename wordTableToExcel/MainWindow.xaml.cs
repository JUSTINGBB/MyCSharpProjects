using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace wordTableToExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void mInputBtn_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            DialogResult folderR = folderDialog.ShowDialog();
            if (folderR == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            String pathStr = folderDialog.SelectedPath.Trim();
            mInputTxt.Text = pathStr;
        }

        private void mOutputBtn_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            DialogResult folderR = folderDialog.ShowDialog();
            if (folderR == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            String pathStr = folderDialog.SelectedPath.Trim();
            mOutputTxt.Text = pathStr;
        }

        //运行
        private void mStartProcess_Click(object sender, RoutedEventArgs e)
        {
            mMessage.Text = "";
            if (mInputTxt.Text == "")
            {
                mMessage.Text = "请选择选择输入文件夹";
            }
            else 
            {
                if (Directory.Exists(@mInputTxt.Text))//文件夹存在
                {
                    GetDocxFile(@mInputTxt.Text);
                } 
            }                      
        }

        //遍历文件夹中的docx文档
        public void GetDocxFile(string dirPath)
        {
            DirectoryInfo d = new DirectoryInfo(dirPath);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    GetDocxFile(fsinfo.FullName);//递归调用
                }
                else
                {
                    //Console.WriteLine(fsinfo.FullName);//输出文件的全部路径
                    if (fsinfo.Extension == ".docx")
                    {
                        ReadWord(fsinfo.FullName);
                    }                   
                }
            }
        }
        
        //读取docx文件提取表格
        public void ReadWord(string fileName)
        {
            if (File.Exists(fileName))
            {
                using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(fileName, true))
                {
                    // Insert other code here
                    Body body = wdDoc.MainDocumentPart.Document.Body;
                   
                }
            }
            else
            { 
                mMessage.Text = "文件不存在";
            }
            
        }

    }
}
