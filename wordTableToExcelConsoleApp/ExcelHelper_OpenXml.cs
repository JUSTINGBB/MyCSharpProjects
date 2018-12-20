using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace functionTest
{
    class ExcelHelper_OpenXml
    {
        //创建excel
        public static void CreateExcel(string filepath)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            //通过文件路径创建一个电子表格document
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            //添加一个WorkbookPart到document
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            //添加一个WorksheetPart到WorkbookPart
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            //Add Sheets to the Workbook.
            //添加sheets到Workbook
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            //添加一个新的worksheet并与workbook关联
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();
            //Close the  document.
            spreadsheetDocument.Close();
        }

        //读取excel的所有sheets名
        public static void ReadExcel(string excelPathStr)
        {
            //检索工作簿中所有工作表的列表。             
            // Sheets类包含一个集合             
            // OpenXmlElement对象，每个对象代表一个表格。
            Sheets theSheets = null;
            using (SpreadsheetDocument excelDoc = SpreadsheetDocument.Open(excelPathStr, true))
            {
                
                WorkbookPart wbPart = excelDoc.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
                foreach (Sheet theSheet in theSheets)
                {
                    Console.WriteLine(theSheet.Name);
                }
            }
        }
        // The DOM approach.
        // Note that the code below works only for cells that contain numeric values.
        // 读取电子表格文本的DOM方法，请注意，下面的代码仅适用于包含数值的单元格。
        public static void ReadExcelFileDOM(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c.CellValue.Text;
                        Console.Write(text + " ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }
        // The SAX approach.
        //读取大型图表SAX方法，防止内存溢出
        public static void ReadExcelFileSAX(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                string text;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(CellValue))
                    {
                        text = reader.GetText();
                        Console.Write(text + " ");
                    }
                }
                Console.WriteLine();
                Console.ReadKey();
            }
        }

        /// <summary>
        /// 插入单个文本到A1
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="text"></param>
        /// <param name="createNewSheet"></param>
        public static void InsertText(string docName, string text,string cellName,bool createNewSheet = true)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                //获取SharedStringTablePart，如果不存在创建一个新的
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                // Insert the text into the SharedStringTablePart.
                //插入文本到SharedStringTablePart
                int index = InsertSharedStringItem(text, shareStringPart);             
                if (createNewSheet)
                {
                    // Insert a new worksheet.
                    //插入新的worksheet
                    WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
                    // Insert cell A1 into the new worksheet.
                    //插入cell(单元格) A1插入新的worksheet
                    Cell cell = InsertCellInWorksheet(cellName, worksheetPart);

                    // Set the value of cell A1.
                    //设置单元格A1的值
                    cell.CellValue = new CellValue(index.ToString());
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                    // Save the new worksheet.
                    //保存新的worksheet
                    worksheetPart.Worksheet.Save();
                }
                else
                {
                     WorksheetPart worksheetPart = spreadSheet.WorkbookPart.WorksheetParts.First();
                     // Insert cell A1 into the new worksheet.
                     //插入cell(单元格) A1插入新的worksheet
                     Cell cell = InsertCellInWorksheet(cellName, worksheetPart);

                     // Set the value of cell A1.
                     //设置单元格A1的值
                     cell.CellValue = new CellValue(index.ToString());
                     cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                     // Save the new worksheet.
                     //保存新的worksheet
                     worksheetPart.Worksheet.Save();
                }
               
            }
        }

        /// <summary>
        /// 插入Table数据到excel
        /// </summary>
        /// <param name="excelName"></param>
        /// <param name="wordTables"></param>
        public static void InsertTables(string excelName, IEnumerable<WP.Table> wordTables)
        {
            using (SpreadsheetDocument spreadSheetDoc = SpreadsheetDocument.Open(excelName, true))
            {
                uint tableIndex = 0;
                uint rowIndex = 0;
                foreach (WP.Table wordTable in wordTables)
                {
                    tableIndex += 1;
                    // Get the SharedStringTablePart. If it does not exist, create a new one.
                    //获取SharedStringTablePart，如果不存在创建一个新的
                    SharedStringTablePart shareStringPart;
                    if (spreadSheetDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        shareStringPart = spreadSheetDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {
                        shareStringPart = spreadSheetDoc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    }
                    //uint rowIndex = 0;
                    foreach(WP.TableRow wTableRow in wordTable.Elements<WP.TableRow>()){
                        rowIndex += 1;
                        int cellIndex = 0;
                        foreach(WP.TableCell wTableCell in wTableRow.Elements<WP.TableCell>()){
                            cellIndex += 1;
                            //插入文本到SharedStringTablePart
                            int index = InsertSharedStringItem(wTableCell.InnerText, shareStringPart);

                            //获取第一个worksheetPart
                            WorksheetPart worksheetPart = spreadSheetDoc.WorkbookPart.WorksheetParts.First();
                            // Insert cell  into the  worksheet.
                            //通过行列号，插入cell(单元格)到worksheet
                            Cell cell = InsertCellInWorksheet(getColumnName(cellIndex)+ rowIndex, worksheetPart);

                            // Set the value of cell .
                            //设置单元格的值
                            cell.CellValue = new CellValue(index.ToString());
                            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                            // Save theworksheet.
                            //保存worksheet
                            worksheetPart.Worksheet.Save();
                            
                        }
                    }

                }
            }
        }
  
        /// <summary>
        /// 给定文本和SharedStringTablePart，创建具有指定文本的SharedStringItem 
        /// 并将其插入SharedStringTablePart。如果该项已存在，则返回其索引。
        /// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        /// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="shareStringPart"></param>
        /// <returns></returns>
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        /// <summary>
        /// Given a WorkbookPart, inserts a new worksheet.
        /// 插入一个新的worksheet到一个WorkbookPart中
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <returns></returns>
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        /// <summary>
        /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        /// If the cell already exists, returns it. 
        /// 给一个列名，一个行号，和一个WorksheetPart，插入一个cell到worksheet
        /// 如果存在，则返回它
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="rowIndex"></param>
        /// <param name="worksheetPart"></param>
        /// <returns></returns>
        private static Cell InsertCellInWorksheet(string cellName, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = cellName;
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            // If the worksheet does not contain a row with the specified row index, insert one.
            //如果worksheet不包含有具体列号的列，插入一个列
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                //根据CellReference，单元格必须按顺序排列。确定插入新单元格的位置
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="docName"></param>
        /// <param name="sheetName"></param>
        /// <param name="cell1Name"></param>
        /// <param name="cell2Name"></param>
        private static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                Worksheet worksheet = GetWorksheet(document, sheetName);
                if (worksheet == null || string.IsNullOrEmpty(cell1Name) || string.IsNullOrEmpty(cell2Name))
                {
                    return;
                }

                // Verify if the specified cells exist, and if they do not exist, create them.
                CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
                CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

                MergeCells mergeCells;
                if (worksheet.Elements<MergeCells>().Count() > 0)
                {
                    mergeCells = worksheet.Elements<MergeCells>().First();
                }
                else
                {
                    mergeCells = new MergeCells();

                    // Insert a MergeCells object into the specified position.
                    if (worksheet.Elements<CustomSheetView>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                    }
                    else if (worksheet.Elements<DataConsolidate>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                    }
                    else if (worksheet.Elements<SortState>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                    }
                    else if (worksheet.Elements<AutoFilter>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                    }
                    else if (worksheet.Elements<Scenarios>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                    }
                    else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                    }
                    else if (worksheet.Elements<SheetProtection>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                    }
                    else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                    }
                    else
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                    }
                }

                // Create the merged cell and append it to the MergeCells collection.
                MergeCell mergeCell = new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
                mergeCells.Append(mergeCell);

                worksheet.Save();
            }
        }
        
        // Given a Worksheet and a cell name, verifies that the specified cell exists.
        // If it does not exist, creates a new cell. 
        private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

            // If the Worksheet does not contain the specified row, create the specified row.
            // Create the specified cell in that row, and insert the row into the Worksheet.
            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Descendants<SheetData>().First().Append(row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                // If the row does not contain the specified cell, create the specified cell.
                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }

        /// <summary>
        /// Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
        /// 通过文件名和sheet名获取对应的worksheet
        /// </summary>
        /// <param name="document"></param>
        /// <param name="worksheetName"></param>
        /// <returns></returns>
        private static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
            if (sheets.Count() == 0)
                return null;
            else
                return worksheetPart.Worksheet;
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// 例如"A1"，获取“A”
        /// </summary>
        /// <param name="cellName"></param>
        /// <returns></returns>
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        //将列号转化为excel列形式“字母A-Z”
        private static string getColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the row index.
        /// 例如"A1"，获取“1”
        /// </summary>
        /// <param name="cellName"></param>
        /// <returns></returns>
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }
    }
}
