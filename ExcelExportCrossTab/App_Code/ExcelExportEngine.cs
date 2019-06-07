using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

public class ExcelExportEngine
{
    private readonly IWorkbook _excelWorkbook;
    private ISheet _sheet;
    private string _controlType = "crosstab";

    public ExcelExportEngine()
    {
        this._excelWorkbook = new XSSFWorkbook();
        _sheet = _excelWorkbook.CreateSheet("demo");
    }
    public void saveTofle(MemoryStream file, string fileName)
    {
        using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write))
        {
            byte[] buffer = file.ToArray();//转化为byte格式存储
            fs.Write(buffer, 0, buffer.Length);
            fs.Flush();
            buffer = null;
        }//使用using可以最后不用关闭fs 比较方便
    }

    public void Close()
    {
        //把编辑过后的工作薄重新保存为excel文件
        FileStream fs2 = File.Create(@"D:\测试.xlsx");
        _excelWorkbook.Write(fs2);
        fs2.Close();
    }

    #region 创建Sheet
    public void CreateSheet()
    {
        if (this._controlType == "crosstab")//交叉表
        {
            // Retrieve the data table from Excel Data Source.
            DataTable dt = ExcelLayer.GetDataTable("DataForPivot.xls", "Sheet1$");
            /*
                Note:: If you wish to read the data from excel, uncomment the above code and comment the below code.//
                */
            //DataTable dt = SqlLayer.GetDataTable("GetEmployee");
            Pivot pvt = new Pivot(dt);
            int rowDimentionCount = 2;//行维度数量
            int columnDimentionCount = 3;//列维度数量
            int valueDimentionCount = 2;//值维度数量
            string[] rowDimensions = {  "Year", "Company" };
            string[] columnDimensions = {  "Department", "Name", "Designation" };
            columnDimentionCount = columnDimensions.Length;
            string[] valueDimensions = { "CTC", "IsActive" };
            valueDimentionCount = valueDimensions.Length;
            bool rowGroup = true; // 行小计
            bool colGroup = true; // 列小计
            bool rowSum = true; // 行合计
            bool colSum = true; // 列合计
            DataTable dtnew = pvt.PivotData(valueDimensions, AggregateFunction.Sum, rowDimensions, columnDimensions, rowGroup, colGroup, rowSum, colSum);

            int columnCount = dtnew.Columns.Count;//列数
            string title = "测试";
            int rowIndex = 0;
            if (title != "")
            {
                //创建标题行
                CreateRow(_sheet, null, rowIndex, columnCount, 20);
                _sheet.GetRow(0).GetCell(0).SetCellValue(title);
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, columnCount - 1);
                _sheet.AddMergedRegion(cellRangeAddress);
                rowIndex += 1;
            }

            string[] rows = dtnew.Columns[rowDimentionCount].ColumnName.Split('.');//标题维度信息
            int rowLength = rows.Length;
            for (int i = 0; i < rowLength; i++)
            {
                CreateRow(_sheet, null, i + rowIndex, columnCount, 200);
            }
            //创建标题行
            for (int j = 0; j < columnCount; j++)
            {
                for (int i = 0; i < rowLength; i++)
                {
                    string[] currentvalue = dtnew.Columns[j].ColumnName.Split('.');
                    if (j < rowDimentionCount)
                    {
                        _sheet.GetRow(i + rowIndex).GetCell(j).SetCellValue(rowDimensions[j]);
                    }
                    else
                    {
                        if (i == rowLength - 1)
                        {
                            int currentcol = (j - rowDimentionCount) % valueDimentionCount;
                            _sheet.GetRow(i + rowIndex).GetCell(j).SetCellValue(valueDimensions[currentcol]);
                        }
                        else
                        {
                            _sheet.GetRow(i + rowIndex).GetCell(j).SetCellValue(currentvalue[i]);
                        }
                    }

                }
                //合并行标题上的行标题数据
                if (j < rowDimentionCount)
                {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex + rowLength - 1, j, j);
                    _sheet.AddMergedRegion(cellRangeAddress);
                }
            }
            int rowCount = dtnew.Rows.Count;//行数
            for (int i = 0; i < rowCount; i++)
            {
                CreateRow(_sheet, null, i + rowLength + rowIndex, columnCount, 200);
            }
            //填充数据
            for (int j = 0; j < columnCount; j++)
            {
                for (int i = 0; i < rowCount; i++)
                {
                    string currentvalue = dtnew.Rows[i][j].ToString();
                    _sheet.GetRow(i + rowLength + rowIndex).GetCell(j).SetCellValue(currentvalue);
                    //行标题样式
                    _sheet.GetRow(i + rowLength + rowIndex).GetCell(j).CellStyle = null;
                }
            }
            //合并单元格
            //合计列总计
            int mergeCol = columnCount;
            if (colSum)
            {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex + rowLength - 2, columnCount - valueDimentionCount, columnCount - 1);
                _sheet.AddMergedRegion(cellRangeAddress);
                mergeCol = columnCount - valueDimentionCount;
            }
            //合并行标题的列标题数据
            MergeRowWorkSheet(rowIndex, 0, rowLength, mergeCol, valueDimentionCount, rowDimentionCount);
            //合计行总计
            if (rowSum)
            {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex + rowLength + rowCount - 1, rowIndex + rowLength + rowCount - 1, 0, rowDimentionCount - 1);
                _sheet.AddMergedRegion(cellRangeAddress);
            }
            //合并列标题的行标题数据
            MergeWorkSheet(rowLength + rowIndex, 0, rowCount, rowDimentionCount);
        }
    }
    #endregion


    #region 创建行
    private void CreateRow(ISheet sheet, ICellStyle style, int rowIndex, int colCount, float rowHeight)
    {
        IRow row = sheet.CreateRow(rowIndex);
        row.HeightInPoints = rowHeight;//UtilConverter.MillimetersToPoints(rowHeight / 10f);
        for (int i = 0; i < colCount; i++)
        {
            row.CreateCell(i);
            if (style != null) row.GetCell(i).CellStyle = style;
        }
    }
    #endregion

    #region 合并单元格
    /// <summary>
    /// 合并单元格
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="firstRow"></param>
    /// <param name="lastRow"></param>
    /// <param name="firstCol"></param>
    /// <param name="lastCol"></param>
    private void SetMergeCell(ISheet sheet, int firstRow, int lastRow, int firstCol, int lastCol)
    {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.AddMergedRegion(cellRangeAddress);
    }
    #endregion

    /// 合并工作表中指定行数和列数数据相同的单元格
    /// </summary>
    /// <param name="sheetIndex">工作表索引</param>
    /// <param name="beginRowIndex">开始行索引</param>
    /// <param name="beginColumnIndex">开始列索引</param>
    /// <param name="rowCount">要合并的行数</param>
    /// <param name="columnCount">要合并的列数</param>
    public void MergeWorkSheet(int beginRowIndex, int beginColumnIndex, int rowCount, int columnCount)
    {

        //检查参数
        if (columnCount < 1 || rowCount < 1)
            return;

        List<int> rowGroup = new List<int>();

        for (int col = 0; col < columnCount; col++)
        {
            int mark = 0;            //标记比较数据中第一条记录位置
            int mergeCount = 1;        //相同记录数，即要合并的行数
            string text = "";

            for (int row = 0; row < rowCount; row++)
            {
                string prvName = "";
                string nextName = "";
                string prvLastColName = "";
                string nextLastColName = "";

                //最后一行不用比较
                if (row + 1 < rowCount)
                {

                    if (col > beginColumnIndex)
                    {
                        prvLastColName = _sheet.GetRow(beginRowIndex + row).Cells[col - 1].ToString();

                        nextLastColName = _sheet.GetRow(beginRowIndex + row + 1).Cells[col - 1].ToString();
                    }
                    else
                    {
                        prvLastColName = nextLastColName;
                    }
                    prvName = _sheet.GetRow(beginRowIndex + row).Cells[col].ToString();

                    if(col==1 && prvName == "小计")
                    {
                        rowGroup.Add(beginRowIndex + row);
                    }

                    nextName = _sheet.GetRow(beginRowIndex + row + 1).Cells[col].ToString();

                    if (prvName == nextName && prvLastColName == nextLastColName)
                    {
                        mergeCount++;

                        if (row == rowCount - 2)
                        {
                            if (mergeCount != 1)
                            {
                                CellRangeAddress cellRangeAddress = new CellRangeAddress(beginRowIndex + mark, beginRowIndex + mark + mergeCount - 1
                                , beginColumnIndex + col, beginColumnIndex + col);
                                _sheet.AddMergedRegion(cellRangeAddress);
                            }
                        }
                    }
                    else
                    {
                        if (mergeCount != 1)
                        {
                            CellRangeAddress cellRangeAddress = new CellRangeAddress(beginRowIndex + mark, beginRowIndex + mark + mergeCount - 1, beginColumnIndex + col
                            , beginColumnIndex + col);
                            _sheet.AddMergedRegion(cellRangeAddress);
                        }
                        mergeCount = 1;
                        mark = row + 1;
                    }

                }
            }
        }

        //合并行标题的小计
        if (columnCount > 2)
        {
            foreach (var row in rowGroup)
            {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row
                                                , 1, columnCount - 1);
                _sheet.AddMergedRegion(cellRangeAddress);
            }
        }
    }

    /// 合并工作表中指定行数和列数数据相同的单元格
    /// </summary>
    /// <param name="beginRowIndex">开始行索引</param>
    /// <param name="beginColumnIndex">开始列索引</param>
    /// <param name="rowCount">要合并的行数</param>
    /// <param name="columnCount">要合并的列数</param>
    /// <param name="valueCount">值列数</param>
    /// <param name="rowHeadeCount">行标题列数</param>
    public void MergeRowWorkSheet(int beginRowIndex, int beginColumnIndex, int rowCount, int columnCount, int valueCount,int rowHeadeCount)
    {

        //检查参数
        if (columnCount < 1 || rowCount < 1)
            return;
        List<int> colGroup = new List<int>();
        for (int row = beginRowIndex; row < beginRowIndex + rowCount; row++)
        {
            int mark = 0;            //标记比较数据中第一条记录位置
            int mergeCount = 1;        //相同记录数，即要合并的行数
            string text = "";
            for (int col = 0; col < columnCount; col++)
            {
                string prvName = "";
                string prvLastRowName = "";
                string nextName = "";
                string nextLastRowName = "";

                if (col + 1 < columnCount)
                {
                    if (row > beginRowIndex)
                    {
                        prvLastRowName = _sheet.GetRow(row - 1).Cells[col].ToString();

                        nextLastRowName = _sheet.GetRow(row - 1).Cells[col + 1].ToString();
                    }
                    else
                    {
                        prvLastRowName = nextLastRowName;
                    }

                    prvName = _sheet.GetRow(row).Cells[col].ToString();

                    if(row == beginRowIndex+1 && (col-rowHeadeCount)%valueCount==0 && prvName == "小计")
                    {
                        colGroup.Add(col);
                    }

                    nextName = _sheet.GetRow(row).Cells[col + 1].ToString();

                    if (prvName == nextName && prvLastRowName == nextLastRowName && (prvName != "小计" || (prvName == "小计" && colGroup.IndexOf(col)<0)))
                    {
                        mergeCount++;

                        if (col == columnCount - 2)//最后两列比较
                        {
                            if (mergeCount != 1)
                            {
                                CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row, beginColumnIndex + mark, beginColumnIndex + mark + mergeCount - 1);
                                _sheet.AddMergedRegion(cellRangeAddress);
                            }
                        }

                    }
                    else
                    {
                        if (mergeCount != 1)
                        {
                            CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row, beginColumnIndex + mark, beginColumnIndex + mark + mergeCount - 1);
                            _sheet.AddMergedRegion(cellRangeAddress);
                        }
                        mergeCount = 1;
                        mark = col + 1;
                    }

                }
            }
        }

        //合并列标题的小计
        if (rowCount > 2)
        {
            foreach(var col in colGroup)
            {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(beginRowIndex+1, beginRowIndex+rowCount-2, col, col+valueCount-1);
               _sheet.AddMergedRegion(cellRangeAddress);
            }
        }
    }
}