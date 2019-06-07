using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

/// <summary>
/// Pivots the data
/// </summary>
public class Pivot
{
    private DataTable _SourceTable = new DataTable();
    private IEnumerable<DataRow> _Source = new List<DataRow>();

    public Pivot(DataTable SourceTable)
    {
        _SourceTable = SourceTable;
        _Source = SourceTable.Rows.Cast<DataRow>();
    }


    public DataTable PivotData(string[] DataFields, AggregateFunction Aggregate, string[] RowFields, string[] ColumnFields, bool rowGroup, bool colGroup, bool rowSum, bool colSum)
    {
        DataTable dt = new DataTable();
        string tempstr = "", comparestr = "";
        int compareIndex = 0;
        string Separator = ".";

        var RowListString = _Source.Select(x => (RowFields.Select(n => x[n]).Aggregate((a, b) => a += Separator + b.ToString())).ToString()).Distinct().OrderBy(m => m).ToList();

        List<string> RowListTemp = new List<string>();

        RowListString.ForEach(i => RowListTemp.Add(i));

        string tempSubAll = "总计";
        if (rowGroup)
        {
            // 最后一行添加小计
            for (int a = 1; a < RowFields.Length; a++)
            {
                foreach (var row in RowListString)
                {
                    string[] temp = row.Split(Separator.ToCharArray());
                    for (int b = 0; b < RowFields.Length; b++)
                    {
                        if (b == 0)
                        {
                            tempstr = temp[b].ToString();
                            comparestr = tempstr;
                            compareIndex = tempstr.Length;
                        }
                        else if (b < a)
                        {
                            tempstr = tempstr + Separator + temp[b].ToString();
                            comparestr = tempstr;
                            compareIndex = tempstr.Length;
                        }
                        else { tempstr = tempstr + Separator + "小计"; }
                    }
                    if (RowListTemp.FindIndex(x => x == tempstr) > 0) continue;
                    int findIndex = RowListTemp.FindLastIndex(x => x.Length > compareIndex && x.IndexOf("小计") < 0 && x.Substring(0, compareIndex) == comparestr);
                    RowListTemp.Insert(findIndex + 1, tempstr);
                }
            }
        }
        if (rowSum)
        {
            //总计
            for (int a = 1; a < RowFields.Length; a++)
            {
                tempSubAll = tempSubAll + Separator + "总计";
            }
            RowListTemp.Insert(RowListTemp.Count, tempSubAll);
        }

        var ColList = _Source.Select(x => (ColumnFields.Select(n => x[n]).Aggregate((a, b) => a += Separator + b.ToString())).ToString()).Distinct().OrderBy(m => m).ToList();

        //dt.Columns.Add(RowFields);
        foreach (string s in RowFields)
            dt.Columns.Add(s);

        List<string> ColListTemp = new List<string>();
        ColList.ForEach(i => ColListTemp.Add(i));
        tempSubAll = "总计";
        if (colGroup)
        {
            // 最后一列添加小计
            for (int a = 1; a < ColumnFields.Length; a++)
            {
                foreach (var col in ColList)
                {
                    string[] temp = col.Split(Separator.ToCharArray());
                    for (int b = 0; b < ColumnFields.Length; b++)
                    {
                        if (b == 0)
                        {
                            tempstr = temp[b];
                            comparestr = tempstr;
                            compareIndex = tempstr.Length;
                        }
                        else if (b < a)
                        {
                            tempstr = tempstr + Separator + temp[b];
                            comparestr = tempstr;
                            compareIndex = tempstr.Length;
                        }
                        else { tempstr = tempstr + Separator + "小计"; }
                    }
                    if (ColListTemp.FindIndex(x => x == tempstr) > 0) continue;
                    int findIndex = ColListTemp.FindLastIndex(x => x.Length > compareIndex && x.IndexOf("小计") < 0 && x.Substring(0, compareIndex) == comparestr);
                    ColListTemp.Insert(findIndex + 1, tempstr);
                }
            }
        }
        if (colSum)
        {
            //总计
            for (int a = 1; a < ColumnFields.Length; a++)
            {
                tempSubAll = tempSubAll + Separator + "总计";
            }
            ColListTemp.Insert(ColListTemp.Count, tempSubAll);
        }

        ColList = new List<string>();
        ColListTemp.ForEach(i => ColList.Add(i));

        foreach (var col in ColList)
        {
            foreach (var dataField in DataFields)
            {
                var allcolumn = col + Separator + dataField;
                dt.Columns.Add(allcolumn);  // Cretes the result columns.//
            }
        }

        foreach (var rowLinshi in RowListTemp)
        {
            DataRow row = dt.NewRow();
            string strFilter = string.Empty;

            string[] rowValues = rowLinshi.Split(Separator.ToCharArray(), StringSplitOptions.None);
            int colIndex = 0;
            foreach (string Field in RowFields)
            {
                row[Field] = rowValues[colIndex];
                if (rowValues[colIndex] != "小计" && rowValues[colIndex] != "总计")
                {
                    strFilter += " and " + Field + " = '" + rowValues[colIndex] + "'";
                }
                colIndex++;
            }
            if (strFilter.Length > 5)
            {
                strFilter = strFilter.Substring(5);
            }
            else
            {
                strFilter += " 1=1 ";
            }

            foreach (var col in ColList)
            {
                foreach (var dataField in DataFields)
                {
                    var allcolumn = col + "." + dataField;

                    string filter = strFilter;
                    string[] strColValues = allcolumn.Split(Separator.ToCharArray(), StringSplitOptions.None);
                    for (int i = 0; i < ColumnFields.Length; i++)
                    {
                        if (strColValues[i] != "小计" && strColValues[i] != "总计")
                        {
                            filter += " and " + ColumnFields[i] + " = '" + strColValues[i] + "'";
                        }
                    }
                    row[allcolumn] = GetData(filter, dataField, Aggregate);
                }
            }
            dt.Rows.Add(row);
        }
        return dt;
    }

    /// <summary>
    /// Retrives the data for matching RowField value and ColumnFields values with Aggregate function applied on them.
    /// </summary>
    /// <param name="Filter">DataTable Filter condition as a string</param>
    /// <param name="DataField">The column name which needs to spread out in Data Part of the Pivoted table</param>
    /// <param name="Aggregate">Enumeration to determine which function to apply to aggregate the data</param>
    /// <returns></returns>
    private object GetData(string Filter, string DataField, AggregateFunction Aggregate)
    {
        try
        {
            DataRow[] FilteredRows = _SourceTable.Select(Filter);
            object[] objList = FilteredRows.Select(x => x[DataField]).ToArray();

            switch (Aggregate)
            {
                case AggregateFunction.Average:
                    return GetAverage(objList);
                case AggregateFunction.Count:
                    return objList.Count();
                case AggregateFunction.Exists:
                    return (objList.Count() == 0) ? "False" : "True";
                case AggregateFunction.First:
                    return GetFirst(objList);
                case AggregateFunction.Last:
                    return GetLast(objList);
                case AggregateFunction.Max:
                    return GetMax(objList);
                case AggregateFunction.Min:
                    return GetMin(objList);
                case AggregateFunction.Sum:
                    return GetSum(objList);
                default:
                    return null;
            }
        }
        catch (Exception ex)
        {
            return "#Error";
        }
    }

    private object GetAverage(object[] objList)
    {
        return objList.Count() == 0 ? null : (object)(Convert.ToDecimal(GetSum(objList)) / objList.Count());
    }
    private object GetSum(object[] objList)
    {
        return objList.Count() == 0 ? null : (object)(objList.Aggregate(new decimal(), (x, y) => x += Convert.ToDecimal(y)));
    }
    private object GetFirst(object[] objList)
    {
        return (objList.Count() == 0) ? null : objList.First();
    }
    private object GetLast(object[] objList)
    {
        return (objList.Count() == 0) ? null : objList.Last();
    }
    private object GetMax(object[] objList)
    {
        return (objList.Count() == 0) ? null : objList.Max();
    }
    private object GetMin(object[] objList)
    {
        return (objList.Count() == 0) ? null : objList.Min();
    }

}

public enum AggregateFunction
{
    Count = 1,
    Sum = 2,
    First = 3,
    Last = 4,
    Average = 5,
    Max = 6,
    Min = 7,
    Exists = 8
}