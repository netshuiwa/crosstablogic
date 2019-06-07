using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.OleDb;

// Written by Anurag Gandhi.
// Url: http://www.gandhisoft.com
// Contact me at: soft.gandhi@gmail.com

/// <summary>
/// Acts as a DataBase Layer
/// </summary>
public class ExcelLayer
{
	public ExcelLayer()
	{
		//
		// TODO: Add constructor logic here
		//
	}
    /// <summary>
    /// Retireves the data from Excel Sheet to a DataTable.
    /// </summary>
    /// <param name="FileName">File Name along with path from the root folder.</param>
    /// <param name="TableName">Name of the Table of the Excel Sheet. Sheet1$ if no table.</param>
    /// <returns></returns>
    public static DataTable GetDataTable(string FileName, string TableName)
    {
        try
        {
            string strPath = AppContext.BaseDirectory + FileName;
            DataSet ds = new DataSet();
            String sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; " + "Data Source=" + strPath + "; " + "Extended Properties=Excel 8.0;";

            OleDbConnection objConn = new OleDbConnection(sConnectionString);
            objConn.Open();
            OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [" + TableName + "] where IsActive = 1", objConn);
            OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();
            objAdapter1.SelectCommand = objCmdSelect;
            objAdapter1.Fill(ds);
            objConn.Close();
            return ds.Tables[0];
        }
        catch (Exception ex)
        {
            //Log your exception here.//
            return (DataTable)null;
        }
    }
}
