using com.data;
using com.data.structure;
using com.tools.db;
using com.tools.file;
using com.tools.json;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Web.Services;

/// <summary>
/// BroadcastService 的摘要说明
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
[System.Web.Script.Services.ScriptService]

public class BroadcastService : System.Web.Services.WebService {

    System.DateTime currentTime = new System.DateTime();
    private string programname = "BroadcastService";
    private string errlog = "";

    public BroadcastService () {

        //如果使用设计的组件，请取消注释以下行 
        //InitializeComponent(); 
    }

    [WebMethod(Description = "查询  返回 DataTable")]
    public DataTable GetBroadcastViewData(string user_no)
    {
        DataTable dt = new DataTable();
        string strSql = "";

        try
        {
            if(!string.IsNullOrEmpty(user_no))
            {
                Info info = new Info();
                Command cmd = new Command(info);

                strSql = strSql + " SELECT  ";
                strSql = strSql + " A.[ID] AS [ID],";
                strSql = strSql + " A.[Title] AS [Title],";
                strSql = strSql + " A.[TypeID] AS [TypeID],";
                strSql = strSql + " dbo.Get_ItemName('BroadcastType',A.[TypeID]) AS [Type],";
                strSql = strSql + " A.[Content] AS [Content],";
                strSql = strSql + " A.[Link] AS [Link],";
                strSql = strSql + " dbo.Get_FormatDate(A.[Start_Date]) AS [Start_Date], ";
                strSql = strSql + " dbo.Get_FormatDate(A.[End_Date]) AS [End_Date], ";
                strSql = strSql + " dbo.Get_FormatDate(A.[Effective_Date]) AS [Effective_Date], ";
                strSql = strSql + " B.[ID] AS [BroadcastUserId], ";
                strSql = strSql + " B.[FROM_ID] AS [FROM_ID],";
                strSql = strSql + " dbo.Get_BroadcastFrom(B.[FROM_ID]) AS [FROM_NAME], ";
                strSql = strSql + " B.[User_NO] AS [User_NO],";
                strSql = strSql + " B.[Status_ID] AS [Status_ID],";
                strSql = strSql + " dbo.Get_ItemName('BroadCast_Status_ID',B.[Status_ID]) AS [Status],";
                strSql = strSql + " ISNULL(B.[Folder_ID],'') AS [Folder_ID],";
                strSql = strSql + " dbo.[Get_BroadcastFolder](B.[User_NO],B.[Folder_ID]) AS [Folder_Name]";
                strSql = strSql + " FROM dbo.[Broadcast] AS A LEFT JOIN dbo.[Broadcast_User] AS B ON A.ID = B.Broadcast_ID";
                strSql = strSql + " WHERE B.User_NO = '" + user_no + "'";            

                dt = cmd.GetDataTable("BroadcastView", strSql);
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    [WebMethod(Description = "查询  返回 DataTable")]
    public DataTable GetBroadcastViewFolder(string user_no)
    {
        DataTable dt = new DataTable();
        string strSql = "";

        try
        {
            if (!string.IsNullOrEmpty(user_no))
            {
                Info info = new Info();
                Command cmd = new Command(info);

                strSql = "";
                strSql = strSql + " SELECT Sys_ItemValue,Sys_ItemName FROM User_Define WHERE Sys_Code = 'BroadCast_Folder_ID' AND User_NO='" + user_no + "' ORDER BY Sys_ItemValue ASC  ";
                dt = cmd.GetDataTable("BroadcastViewFolder", strSql);

                // 如果使用者没有自定义资料夹 就用系统预设的
                if (dt.Rows.Count == 0) 
                {
                    strSql = "";
                    strSql = strSql + " SELECT Sys_ItemValue,Sys_ItemName FROM SystemCode WHERE Sys_Code = 'BroadCast_Folder_ID' ORDER BY Sys_ItemValue ASC  ";
                    dt = cmd.GetDataTable("BroadcastViewFolder", strSql);
                }
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    [WebMethod(Description = "新增 Broadcast ，返回受影响行数")]
    public int SetBroadcastViewData(byte[] _broadcast)
    {
        int result = 0;
        string sql = "";

        try
        {
            Broadcast broadcastviewdata = new Broadcast();
            broadcastviewdata = (Broadcast)Serialize.DeserializeObject(_broadcast);

            sql = sql + " UPDATE dbo.[Broadcast_User] Set";

            //Vesper 按了连结 才代表读取
            if (broadcastviewdata.Status_ID != "1")  
            { 
                sql = sql + " [Status_ID] = ?,";
                sql = sql + " [Read_DateTime] = ?,";
            }
            sql = sql + " [Folder_ID] = ?";
            sql = sql + " WHERE  ID = ?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();

            if (broadcastviewdata.Status_ID != "1")
            {
                OleDbParameters.Add(new OleDbParameter("@Status_ID", broadcastviewdata.Status_ID));
                OleDbParameters.Add(new OleDbParameter("@Read_DateTime", System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss")));
            }
            OleDbParameters.Add(new OleDbParameter("@Folder_ID", broadcastviewdata.Folder_ID));
            OleDbParameters.Add(new OleDbParameter("@ID", broadcastviewdata.BroadcastUserId));

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


    [WebMethod(Description = "查询  返回 DataTable")]
    public DataTable GetBroadcastData(string input = "")
    {
        DataTable dt = new DataTable();
        string strSql = "";
        string[] data = input.Split(new Char[] { ',' });
        string id = data[0].Trim();

        try
        {
            Info info = new Info();
            Command cmd = new Command(info);

            strSql = strSql + " SELECT "; 
            strSql = strSql + " [ID],"; 
            //VESPER 发布与否 20150115			
            strSql = strSql + "(CASE WHEN (SELECT COUNT(ID) FROM [dbo].[Broadcast_User] WHERE Broadcast_ID = [Broadcast].ID) > 0 THEN ";
			strSql = strSql + "		'Y' ";
			strSql = strSql + " ELSE ";
			strSql = strSql + " 	'N'";
            strSql = strSql + " END) AS POST,";
            
            strSql = strSql + " [Title],"; 
            strSql = strSql + " [TypeID],";
            strSql = strSql + " (dbo.Get_ItemName('BroadcastType',[Broadcast].TypeID)) AS [Type],"; 
            strSql = strSql + " [Content],"; 
            strSql = strSql + " [Link],"; 
            strSql = strSql + " (SUBSTRING([Start_Date],1,4) + '-' + SUBSTRING([Start_Date],5,2) + '-' + SUBSTRING([Start_Date],7,2)) AS [Start_Date],"; 
            strSql = strSql + " (SUBSTRING([End_Date],1,4) + '-' + SUBSTRING([End_Date],5,2) + '-' + SUBSTRING([End_Date],7,2)) AS [End_Date],"; 
            strSql = strSql + " (SUBSTRING([Effective_Date],1,4) + '-' + SUBSTRING([Effective_Date],5,2) + '-' + SUBSTRING([Effective_Date],7,2)) AS [Effective_Date],"; 
            strSql = strSql + " [Target], "; 
            strSql = strSql + " [Organization] ";
            strSql = strSql + " FROM [Broadcast] ";
            strSql = strSql + " WHERE 1=1 ";
            if (!string.IsNullOrEmpty(id))
            {
                strSql = strSql + " AND [Broadcast].ID ='" + id + "' ";
            }
            dt = cmd.GetDataTable("Broadcast", strSql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    [WebMethod(Description = "新增 Broadcast ，返回受影响行数")]
    public int InsertBroadcastData(byte[] _broadcast)
    {
        int result = 0;
        string sql = "";

        try
        {
            Broadcast broadcast = new Broadcast();
            broadcast = (Broadcast)Serialize.DeserializeObject(_broadcast);

            sql = "";
            sql = sql + " INSERT INTO dbo.[Broadcast]([Title],[Content],[Link],[Start_Date],[End_Date],[Effective_Date],[Organization],[Target])";
            sql = sql + " VALUES(?,?,?,?,?,?,?,?)";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Title", broadcast.Title));
            OleDbParameters.Add(new OleDbParameter("@Content", broadcast.Content));
            OleDbParameters.Add(new OleDbParameter("@Link", broadcast.Link));
            OleDbParameters.Add(new OleDbParameter("@Start_Date", broadcast.Start_Date));
            OleDbParameters.Add(new OleDbParameter("@End_Date", broadcast.End_Date));
            OleDbParameters.Add(new OleDbParameter("@Effective_Date", broadcast.Effective_Date));
            OleDbParameters.Add(new OleDbParameter("@Organization", broadcast.Organization));
            OleDbParameters.Add(new OleDbParameter("@Target", broadcast.Target));

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
    
    [WebMethod(Description = "更新 Broadcast ，返回受影响行数")]
    public int UpdateBroadcastData(byte[] _broadcast)
    {
        int result = 0;
        string sql = "";

        try
        {
            Broadcast broadcast = new Broadcast();
            broadcast = (Broadcast)Serialize.DeserializeObject(_broadcast);

            sql = sql + " UPDATE dbo.[Broadcast] Set";
            sql = sql + " [Title] = ?,";
            sql = sql + " [Content] = ?,";
            sql = sql + " [Link] = ?,";
            sql = sql + " [Start_Date] = ?,";
            sql = sql + " [End_Date] = ?,";
            sql = sql + " [Effective_Date] = ?,";
            sql = sql + " [Organization] = ?,";
            sql = sql + " [Target] = ?";
            sql = sql + " WHERE  ID = ?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Title", broadcast.Title));
            OleDbParameters.Add(new OleDbParameter("@Content", broadcast.Content));
            OleDbParameters.Add(new OleDbParameter("@Link", broadcast.Link));
            OleDbParameters.Add(new OleDbParameter("@Start_Date", broadcast.Start_Date));
            OleDbParameters.Add(new OleDbParameter("@End_Date", broadcast.End_Date));
            OleDbParameters.Add(new OleDbParameter("@Effective_Date", broadcast.Effective_Date));
            OleDbParameters.Add(new OleDbParameter("@Organization", broadcast.Organization));
            OleDbParameters.Add(new OleDbParameter("@Target", broadcast.Target));
            OleDbParameters.Add(new OleDbParameter("@ID", broadcast.ID));

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

    [WebMethod(Description = "删除 Broadcast ，返回受影响行数")]
    public int DeleteBroadcastData(string id)
    {
        int result = 0;
        string sql = "";

        try
        {
            if (!string.IsNullOrEmpty(id.Trim()))
            {
                sql = "";
                sql = sql + " DELETE FROM dbo.[Broadcast] WHERE ID='" + id.Trim() + "' ";
                Info info = new Info();
                Command cmd = new Command(info);
                result = cmd.Execute(sql);
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

    [WebMethod(Description = "发布 Broadcast ，返回受影响行数")]
    public int PostBroadcast(string input)
    {
        int result = 0;
        string strSQL = "";

        try
        {
            strSQL = strSQL + "exec SP_Post_Broadcast " + input;

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.Execute(strSQL);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return result;
    }

    [WebMethod(Description = "取销发布 Broadcast ，返回受影响行数")]
    public int UnPostBroadcast(string input)
    {
        int result = 0;
        string strSQL = "";

        try
        {
            strSQL = strSQL + " DELETE FROM [dbo].[Broadcast_User] WHERE Broadcast_ID = " + input;

            Info info = new Info();
            Command cmd = new Command(info);
            result = cmd.Execute(strSQL);
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
