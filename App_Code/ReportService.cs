using com.tools.db;
using com.tools.file;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.ServiceModel.Activation;

[AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]

// 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码、svc 和配置文件中的类名“ReportService”。
public class ReportService : IReportService
{
    System.DateTime currentTime = new System.DateTime();
    private string programname = "ReportService";
    private string errlog = "";

    public DataTable getRpt0001(string input = "") 
    {
        DataTable dt = new DataTable();
        string sql = "";
        string[] Branchs = input.Split(new Char[] { ',' });

        try
        {
            Info info = new Info();
            Command cmd = new Command(info);
            sql = sql + " SELECT * FROM [dbo].[View_Rpt00001] ";

            if (!string.IsNullOrEmpty(input)) 
            {
                sql = sql + "WHERE BRANCH IN (";
                for (int i = 0; i <= Branchs.Length - 1; i++)
                {
                    sql = sql + "'" + Branchs[i].ToString() + "',";
                }
                sql = sql.Substring(0, sql.Length - 1);
                sql = sql + ")";
            }
            dt = cmd.GetDataTable("Rpt00001", sql);

        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }


    public int InsertSchool_Students_Status_Count(string input)
    {
        int result = 0;
        string sql = "";
        string[] UserData = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " DELETE FROM [School_Students_Status_Count] WHERE Branch ='" + UserData[0].Trim() + "' AND YearMonth ='" + UserData[2].Trim() + "' AND CatCode = '" + UserData[3].Trim() + "' ";
            sql = sql + " INSERT INTO School_Students_Status_Count(Branch,ReportStatus,YearMonth,CatCode,EndOfYearInReading,Target,Abnormal,InRegistered,InReading,ClassCount,AddStudent,LostStudent,KeepStudent,StayOpenClass,AVGClassStudent)";
            sql = sql + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Branch", UserData[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@ReportStatus", UserData[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@YearMonth", UserData[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@CatCode", UserData[3].Trim()));
            OleDbParameters.Add(new OleDbParameter("@EndOfYearInReading", UserData[4].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Target", UserData[5].Trim()));
            OleDbParameters.Add(new OleDbParameter("@RadAbnormal", UserData[6].Trim()));
            OleDbParameters.Add(new OleDbParameter("@InRegistered", UserData[7].Trim()));
            OleDbParameters.Add(new OleDbParameter("@InReading", UserData[8].Trim()));
            OleDbParameters.Add(new OleDbParameter("@ClassCount", UserData[9].Trim()));
            OleDbParameters.Add(new OleDbParameter("@AddStudent", UserData[10].Trim()));
            OleDbParameters.Add(new OleDbParameter("@LostStudent", UserData[11].Trim()));
            OleDbParameters.Add(new OleDbParameter("@KeepStudent", UserData[12].Trim()));
            OleDbParameters.Add(new OleDbParameter("@StayOpenClass", UserData[13].Trim()));
            OleDbParameters.Add(new OleDbParameter("@AVGClassStudent", UserData[14].Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.ExecuteSQL(sql, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    public DataTable getSchool_Students_Status_Count(string input)
    {
        DataTable dt = new DataTable();
        string sql = "";
        string[] Conditions = input.Split(new Char[] { ',' });
        string strBranch_ID = Conditions[0].Trim();
        string strYearMonth = Conditions[1].Trim();
        string strCatCode = Conditions[2].Trim();

        try
        {
            sql = sql + " SELECT * ";
            sql = sql + " FROM [School_Students_Status_Count]";
            sql = sql + " WHERE 1=1 ";
            if (!string.IsNullOrEmpty(strBranch_ID))
            {
                sql = sql + "AND Branch = '" + strBranch_ID + "' ";
            }
            if (!string.IsNullOrEmpty(strYearMonth))
            {
                sql = sql + "AND YearMonth = '" + strYearMonth + "' ";
            }
            if (!string.IsNullOrEmpty(strCatCode))
            {
                sql = sql + "AND CatCode = '" + strCatCode + "' ";
            }
            sql = sql + " Order By Branch,YearMonth,CatCode ASC";

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("School_Students_Status_Count", sql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public int setReportStatus(string input) 
    {
        int result = 0;
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });

        try
        {
            sql = sql + " UPDATE [dbo].[Rpt00001] SET";
            sql = sql + "  ReportStatus = ?,Mdy_UserID = ?,Mdy_DateTime = ?";
            sql = sql + " WHERE ID = ?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@ReportStatus", data[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Mdy_UserID", data[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Mdy_DateTime",  System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss")));
            OleDbParameters.Add(new OleDbParameter("@ID", data[2].Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.ExecuteSQL(sql,OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    /// <summary>
    ///    判断 Rpt00001 是否有资料
    /// </summary>
    /// <param name="input">条件字串流</param>
    /// <returns>ture/false</returns>
    public bool chkRpt00001(string input) 
    {
        bool result = false;
        string sql = "";
        string[] Conditions = input.Split(new Char[] { ',' });
        string branch = Conditions[0].Trim();
        string catcode = Conditions[1].Trim();
        string firstdayofweek = Conditions[2].Trim();
        string lastdayofweek = Conditions[3].Trim();
        int rowcount = 0;

        try
        {
            sql = sql + " SELECT COUNT(ID) AS CNT ";
            sql = sql + " FROM [dbo].[Rpt00001] ";
            sql = sql + " WHERE ";
            sql = sql + " Branch = '" + branch + "' AND ";
            sql = sql + " CatCode = '" + catcode + "' AND ";
            sql = sql + " (BeginDate = '" + firstdayofweek + "' AND EndDate = '" + lastdayofweek + "')";

            Info info = new Info();
            Command cmd = new Command(info);
            DataTable dt = new DataTable();
            dt = cmd.GetDataTable("Rpt00001", sql);
            rowcount = (int)dt.Rows[0]["CNT"];       
            
            if (rowcount == 1) 
            {
                result = true;
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
   return result;
    }

    public int InsertRpt00001(string input)
    {
        int result = 0;
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " INSERT INTO [dbo].[Rpt00001](Branch,CatCode,BeginDate,EndDate,Target,Abnormal,InRegistered,InReading,ClassCount,AddStudent,LostStudent,KeepStudent,StayOpenClass,AVGClassStudent,ReportStatus,Ins_UserID,Ins_DateTime)";
            sql = sql + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Branch", data[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@CatCode", data[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@BeginDate", data[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@EndDate", data[3].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Target", data[4].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Abnormal", data[5].Trim()));
            OleDbParameters.Add(new OleDbParameter("@InRegistered", data[6].Trim()));
            OleDbParameters.Add(new OleDbParameter("@InReading", data[7].Trim()));
            OleDbParameters.Add(new OleDbParameter("@ClassCount", data[8].Trim()));
            OleDbParameters.Add(new OleDbParameter("@AddStudent", data[9].Trim()));
            OleDbParameters.Add(new OleDbParameter("@LostStudent", data[10].Trim()));
            OleDbParameters.Add(new OleDbParameter("@KeepStudent", data[11].Trim()));
            OleDbParameters.Add(new OleDbParameter("@StayOpenClass", data[12].Trim()));
            OleDbParameters.Add(new OleDbParameter("@AVGClassStudent", data[13].Trim()));
            OleDbParameters.Add(new OleDbParameter("@ReportStatus", data[14].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Ins_UserID", data[15].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Ins_DateTime", System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss")));
            
            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.ExecuteSQL(sql, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    public int UpdateRpt00001(string input)
    {
        int result = 0;
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " UPDATE [dbo].[Rpt00001] SET ";
            sql = sql + "   Target=?,Abnormal=?,InRegistered=?,InReading=?,";
            sql = sql + "   ClassCount=?,AddStudent=?,LostStudent=?,KeepStudent=?,";
            sql = sql + "   StayOpenClass=?,AVGClassStudent=?,ReportStatus=?,Mdy_UserID=?,Mdy_DateTime=? ";
            sql = sql + " WHERE ID=?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Target", data[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Abnormal", data[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@InRegistered", data[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@InReading", data[3].Trim()));
            OleDbParameters.Add(new OleDbParameter("@ClassCount", data[4].Trim()));
            OleDbParameters.Add(new OleDbParameter("@AddStudent", data[5].Trim()));
            OleDbParameters.Add(new OleDbParameter("@LostStudent", data[6].Trim()));
            OleDbParameters.Add(new OleDbParameter("@KeepStudent", data[7].Trim()));
            OleDbParameters.Add(new OleDbParameter("@StayOpenClass", data[8].Trim()));
            OleDbParameters.Add(new OleDbParameter("@AVGClassStudent", data[9].Trim()));
            OleDbParameters.Add(new OleDbParameter("@ReportStatus", data[10].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Mdy_UserID", data[11].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Mdy_DateTime", System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss")));
            OleDbParameters.Add(new OleDbParameter("@ID", data[12].Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.ExecuteSQL(sql, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    public int DeleteRpt00001(string input)
    {
        int result = 0;
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " DELETE FROM [dbo].[Rpt00001] WHERE ID=?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@ID", data[0].Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.ExecuteSQL(sql, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    public int InsertReportUpload(string input)
    {
        int result = 0;
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " INSERT INTO [dbo].[ReportUpload](RU_Branch,RU_BeginDate,RU_EndDate,RU_FileName,Ins_UserID,Ins_DateTime)";
            sql = sql + " VALUES(?,?,?,?,?,?)";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@RU_Branch", data[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@RU_BeginDate", data[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@RU_EndDate", data[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@RU_FileName", data[3].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Ins_UserID", data[4].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Ins_DateTime", System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss")));

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.ExecuteSQL(sql, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    public bool chkReportUpload(string input)
    {
        bool result = false;
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });
        string branch = data[0].Trim();
        string begindate = data[1].Trim();
        string enddate = data[2].Trim();

        try
        {
            sql = sql + " SELECT COUNT(ID) AS CNT ";
            sql = sql + " FROM [dbo].[ReportUpload] ";
            sql = sql + " WHERE ";
            sql = sql + " RU_Branch = '" + branch + "' AND ";
            sql = sql + " RU_BeginDate = '" + begindate + "' AND ";
            sql = sql + " RU_EndDate = '" + enddate + "' ";

            Info info = new Info();
            Command cmd = new Command(info);
            DataTable dt = new DataTable();
            dt = cmd.GetDataTable("ReportUpload", sql);

            if ((int)dt.Rows[0]["CNT"] > 0)
            {
                result = true;
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    public DataTable getReportUpload(string input)
    {
        DataTable dt = new DataTable();
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });
        string branch = data[0].Trim();
        string startdate = data[1].Trim();
        string enddate = data[2].Trim();

        try
        {
            sql = sql + " SELECT * ";
            sql = sql + " FROM [dbo].[View_ReportUpload]";
            sql = sql + " WHERE 1=1 ";

            if (!string.IsNullOrEmpty(branch))
            {
                if (branch.IndexOf("|") > 0)
                {
                    string[] branchs = branch.Split(new Char[] { '|' });

                    sql = sql + "AND Branch_ID IN(";
                    for (int i = 0; i <= branchs.Length - 1; i++)
                    {
                        if (i == branchs.Length - 1)
                        {
                            sql = sql + "'" + branchs[i].Trim() + "'";
                        }
                        else
                        {
                            sql = sql + "'" + branchs[i].Trim() + "',";
                        }
                    }
                    sql = sql + ")";
                }
                else 
                {
                    sql = sql + "AND Branch_ID = '" + branch + "' ";
                }
            }
            if (!string.IsNullOrEmpty(startdate))
            {
                sql = sql + "AND BeginDate = '" + startdate + "' ";
            }
            if (!string.IsNullOrEmpty(enddate))
            {
                sql = sql + "AND EndDate = '" + enddate + "' ";
            }

            sql = sql + " Order By Branch,BeginDate,[FileName] ASC";

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("View_ReportUpload", sql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public int DeleteReportUpload(string input)
    {
        int result = 0;
        string sql = "";
        string[] data = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " DELETE FROM [dbo].[View_ReportUpload] WHERE 1=1 ";

            if (data.Length > 0) 
            {
                sql = sql + "AND ID IN(";
                for (int i = 0; i <= data.Length - 1; i++) 
                {
                    if (i == data.Length - 1)
                    {
                        sql = sql + "'" + data[i].Trim() + "'";
                    }
                    else 
                    {
                        sql = sql + "'" + data[i].Trim() + "',";                
                    }
                }
                sql = sql + ")";
            }

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.Execute(sql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }
}
