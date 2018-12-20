using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;




namespace functionTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //for (; ; )//无限循环
            //{
            //    GetDocxFile(@"E:\国土局相关\历史文档");
            //}
            //GetDocxFile(@"E:\国土局相关\历史文档");
            //Director(@"E:\国土局相关");
            //ReadDocxTables(@"D:\VSprojects\tableTest.docx");
            //AddTable(@"D:\VSprojects\tableTest.docx",new string[,] 
            //    { { "Texas", "TX" }, 
            //    { null , "CA" }, 
            //    { "New York", "NY" }, 
            //    { "Massachusetts", "MA" } }
            //    );
            //ReadDocxTables(@"D:\VSprojects\tableTest.docx");
            //OfficeInteropTest.CreateDocx();
            //Console.WriteLine("点击任意键继续！");
            //Console.ReadLine();
            //Console.ReadKey();
            //Console.Write(Console.ReadKey().Key);

            string wordDocName = @"D:\tableTest.docx";
            string fileName = @"D:\factories.png";
            string excelFileName = @"D://mexcelTest.xlsx";
            //if (!FileIsUsed(fileName))
            //{
            //    Console.WriteLine("文件已被其他程序占用");
            //    Console.ReadKey();
            //    return;
            //}

            ExcelHelper_OpenXml.CreateExcel(@"D://mexcelTest.xlsx");//创建电子表格
            Console.WriteLine("创建电子表格D://mexcelTest.xlsx成功！");
            //ExcelHelper_OpenXml.InsertText(@"D://mexcelTest.xlsx", "你好","A1");//A1插入一个文本

            //创建 docx文档
            //string cDoc = @"D:/CreateAndAddCharacterStyle.docx";
            //WordHelper_OpenXml.CreateWordprocessingDocument(cDoc);

            //读取word表格文本
            ReadDocxTablesText(wordDocName);
            Console.ReadKey();

            //读取word表格，将表格插入excel
            IEnumerable<Table> tables = ReadDocxTableList(wordDocName);//or var tables =****;
            ExcelHelper_OpenXml.InsertTables(excelFileName, tables);
            Console.WriteLine("word表格文本插入excel成功！");
           
            
            //word插入图片
            //WordHelper_OpenXml.InsertAPicture(wordDocName, fileName);

            Console.ReadKey();
        }
        
        //遍历文件夹中的docx文档
        public static string GetDocxFile(string dirPath)
        {
            string fileStr = null;
            DirectoryInfo theFolder = new DirectoryInfo(dirPath);
            FileInfo[] fileInfo = theFolder.GetFiles();
            foreach (FileInfo NextFile in fileInfo)
            {
                if (NextFile.Extension == ".docx")
                {
                    Console.WriteLine(NextFile.Name);
                    fileStr = NextFile.Name;
                }
                
            }
            return fileStr;
        }
        
        //遍历所有文件
        public static  void Director(string dir) 
        {
            DirectoryInfo d = new DirectoryInfo(dir);
            FileSystemInfo[] fsinfos = d.GetFileSystemInfos();
            foreach (FileSystemInfo fsinfo in fsinfos)
            {
                if (fsinfo is DirectoryInfo)     //判断是否为文件夹
                {
                    Director(fsinfo.FullName);//递归调用
                }
                else 
                {
                    Console.WriteLine(fsinfo.FullName);//输出文件的全部路径
                }
            }
               
            
        }

        /// <summary>
        /// 返回指示文件是否已被其它程序使用的布尔值
        /// </summary>
        /// <param name="fileFullName">文件的完全限定名，例如：“C:\MyFile.txt”。</param>
        /// <returns>如果文件已被其它程序使用，则为 true；否则为 false。</returns>
        public static Boolean FileIsUsed(String fileFullName)
        {
            Boolean result = false;
            //判断文件是否存在，如果不存在，直接返回 false
            if (!System.IO.File.Exists(fileFullName))
            {
                result = false;
            }//end: 如果文件不存在的处理逻辑
            else
            {//如果文件存在，则继续判断文件是否已被其它程序使用
                //逻辑：尝试执行打开文件的操作，如果文件已经被其它程序使用，则打开失败，抛出异常，根据此类异常可以判断文件是否已被其它程序使用。
                System.IO.FileStream fileStream = null;
                try
                {
                    fileStream = System.IO.File.Open(fileFullName, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None);
                    result = false;
                }
                catch (System.IO.IOException ioEx)
                {
                    result = true;
                }
                catch (System.Exception ex)
                {
                    result = true;
                }
                finally
                {
                    if (fileStream != null)
                    {
                        fileStream.Close();
                    }
                }
            }//end: 如果文件存在的处理逻辑
            //返回指示文件是否已被其它程序使用的值
            return result;
        }//end method FileIsUsed

        /// <summary>
        /// 读取word文档表格文本,（读取全部文本）
        /// </summary>
        /// <param name="wordPathStr"></param>
        static void ReadDocxTablesText(string wordPathStr)
        {
            //string wordPathStr = @"D:\VSprojects\tableTest.docx";
            try
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(wordPathStr, true))
                {

                    Body body = doc.MainDocumentPart.Document.Body;
                    foreach (var table in body.Elements<Table>())
                    {
                        foreach (var tableRow in table.Elements<TableRow>())
                        {
                            //Console.WriteLine(tableRow.InnerText);//得到的文本一样
                            Console.Write("\n");
                            foreach (var tableCell in tableRow.Elements<TableCell>())
                            {
                                Console.Write(tableCell.InnerText + "|"); var v = tableCell.Elements<VerticalMerge>().FirstOrDefault();
                            }
                        }
                    }
                    //读取全部文本
                    var tableCellList = body.Elements<OpenXmlElement>();
                    foreach (var tableCell in tableCellList)
                    {
                        Console.WriteLine(tableCell.InnerText + "|");
                    }
                    Console.WriteLine("读取word表格文本成功！");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine("读取word表格文本失败！");
            }
        }
        
        //返回word文档所有Table
        static IEnumerable<Table> ReadDocxTableList(string wordPathStr)
        {
            IEnumerable<Table> tables;
            //string wordPathStr = @"D:\VSprojects\tableTest.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(wordPathStr, true))
            {
                
                Body body = doc.MainDocumentPart.Document.Body;
                //List<Table> tableList = new List<Table>();
                IEnumerable<Table> tables1 = body.Elements<Table>();
                //foreach (var table in body.Elements<Table>())
                //{
                //    tableList.Add(table);
                //}
                tables = tables1;
            }
            return tables;
        }
        
        //word 添加表格
        public static void AddTable(string fileName, string[,] data)
        {
            using (var document = WordprocessingDocument.Open(fileName, true))
            {

                var doc = document.MainDocumentPart.Document;

                Table table = new Table();

                TableProperties props = new TableProperties(
                    new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    }));

                table.AppendChild<TableProperties>(props);

                for (var i = 0; i <= data.GetUpperBound(0); i++)
                {
                    var tr = new TableRow();
                    for (var j = 0; j <= data.GetUpperBound(1); j++)
                    {
                        var tc = new TableCell();
                        tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                        if (i == 0 & j == 0)//测试垂直合并单元格，起始点
                        {
                            tc.Append(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Auto },
                            new VerticalMerge { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) },
                            new TableCellVerticalAlignment { Val = new EnumValue<TableVerticalAlignmentValues>(TableVerticalAlignmentValues.Center)}
                            ));
                        }
                        else if (i == 1 & j == 0)//测试垂直合并单元格，终止点
                        {
                            tc.Append(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Auto },
                            new VerticalMerge { },
                            new TableCellVerticalAlignment { Val = new EnumValue<TableVerticalAlignmentValues>(TableVerticalAlignmentValues.Center)}
                            ));
                        }
                        else
                        {
                            // Assume you want columns that are automatically sized.
                            tc.Append(new TableCellProperties(
                                new TableCellWidth { Type = TableWidthUnitValues.Auto }
                                ));
                        }
                        

                        tr.Append(tc);
                    }
                    table.Append(tr);
                    
                }
                doc.Body.Append(table);

                // Add new text.
                //Paragraph para = doc.Body.AppendChild(new Paragraph());
                //Run run = para.AppendChild(new Run());
                //run.AppendChild(new Text("\n"));
                doc.Body.Append(new Paragraph(new Run(new Text("\n"))));

                doc.Save();
            }
        }
              
    }
}
