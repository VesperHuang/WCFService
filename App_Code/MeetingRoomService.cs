using com.data;
using com.data.structure;
using com.tools.db;
using com.tools.file;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.ServiceModel.Activation;


[AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
// 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码、svc 和配置文件中的类名“MeetingRoomService”。
public class MeetingRoomService : IMeetingRoomService
{
    private const string programname = "MeetingRoomService";
    private string errlog = "";

    public List<AppointmentInfo> GetAppointments()
    {
        List<AppointmentInfo> list = new List<AppointmentInfo>();
        DataTable dt = new DataTable();
        string strSql = "";

        try
        {
            Info info = new Info();
            Command cmd = new Command(info);

            strSql = "";
            strSql = strSql + " SELECT";
            strSql = strSql + "     A.[ID] AS [ID],";
            strSql = strSql + " 	A.[Subject] AS [Subject],";
            strSql = strSql + "     A.[Description] AS [Description],";
            strSql = strSql + "     A.[Start] AS [Start],";
            strSql = strSql + "     A.[End] AS [End],";
            strSql = strSql + "     A.[User] AS [User],";
            strSql = strSql + "     A.[Room] AS [Room]";
            strSql = strSql + " FROM [dbo].[Appointment] AS A";
            strSql = strSql + " WHERE 1=1 ";

            //if (!string.IsNullOrEmpty(input))
            //{
            //    strSql = strSql + " AND A.[User_NO] LIKE '%" + input + "%'";
            //}
            strSql = strSql + " ORDER BY A.[ID] ASC";
            dt = cmd.GetDataTable("Appointments", strSql);

            foreach (DataRow row in dt.Rows)
            {
                AppointmentInfo apt = new AppointmentInfo();
                apt.ID = row["ID"].ToString();
                apt.Subject = row["Subject"].ToString();
                apt.Description = row["Description"].ToString();
                apt.Start = (DateTime)row["Start"];
                apt.End = (DateTime)row["End"];
                apt.User = row["User"].ToString();
                apt.Room = row["Room"].ToString();
                list.Add(apt);
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return list;
    }

    public int insertAppointment(AppointmentInfo app) 
    {
        int result = 0;
        string sql = "";

        try
        {
            sql = "";
            sql = sql + " INSERT INTO Appointment([Subject],[Description],[Start],[End],[User],[Room])";
            sql = sql + " VALUES(?,?,?,?,?,?)";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Subject", app.Subject));
            OleDbParameters.Add(new OleDbParameter("@Description", app.Description));
            OleDbParameters.Add(new OleDbParameter("@Start", app.Start));
            OleDbParameters.Add(new OleDbParameter("@End", app.End));
            OleDbParameters.Add(new OleDbParameter("@User", app.User));
            OleDbParameters.Add(new OleDbParameter("@Room", app.Room));

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

    public int updateAppointment(AppointmentInfo app)
    {
        int result = 0;
        string sql = "";

        try
        {
            sql = sql + " UPDATE dbo.[Appointment] Set";
            sql = sql + "   [Subject] = ?,";
            sql = sql + "   [Description] = ?,";
            sql = sql + "   [Start] = ?,";
            sql = sql + "   [End] = ?,";
            sql = sql + "   [User] = ?,";
            sql = sql + "   [Room] = ?";
            sql = sql + " WHERE  ID = ?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Subject", app.Subject));
            OleDbParameters.Add(new OleDbParameter("@Description", app.Description));
            OleDbParameters.Add(new OleDbParameter("@Start", app.Start));
            OleDbParameters.Add(new OleDbParameter("@End", app.End));
            OleDbParameters.Add(new OleDbParameter("@User", app.User));
            OleDbParameters.Add(new OleDbParameter("@Room", app.Room));
            OleDbParameters.Add(new OleDbParameter("@ID", app.ID));

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

    public int deleteAppointment(string id) 
    {
        int result = 0;
        string sql = "";

        try
        {
            sql = "";
            sql = sql + " DELETE FROM [dbo].[Appointment] WHERE ID=?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@ID", int.Parse(id)));

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



    public List<RoomInfo> GetMeetingRooms() 
    {
        List<RoomInfo> list = new List<RoomInfo>();
        DataTable dt = new DataTable();
        string strSql = "";

        try
        {
            Info info = new Info();
            Command cmd = new Command(info);

            strSql = "";
            strSql = strSql + " SELECT";
            strSql = strSql + " 	[Sys_ItemValue],";
            strSql = strSql + " 	[Sys_ItemName]";
            strSql = strSql + " FROM [dbo].[SystemCode]";
            strSql = strSql + " WHERE Sys_Code ='Meeting_Room'";
            strSql = strSql + " ORDER BY [Sys_ItemValue] ASC";
            dt = cmd.GetDataTable("MeetingRooms", strSql);

            foreach (DataRow row in dt.Rows)
            {
                RoomInfo room = new RoomInfo();
                room.RoomNo = row["Sys_ItemValue"].ToString();
                room.RoomName = row["Sys_ItemName"].ToString();
                list.Add(room);
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return list;
    }

    //Vesper this function need to move service
    public List<UserInfo> GetUser()
    {
        List<UserInfo> list = new List<UserInfo>();
        DataTable dt = new DataTable();
        string strSql = "";

        try
        {
            Info info = new Info();
            Command cmd = new Command(info);

            strSql = "";
            strSql = strSql + " SELECT";
            strSql = strSql + " 	[User_NO],";
            strSql = strSql + " 	([User_CnName] + '(' + [User_EnName] + ')' ) AS [User_Name]";
            strSql = strSql + " FROM [dbo].[UserTable]";
            strSql = strSql + " WHERE [User_Type] = '2' AND ISNULL([User_GroupID],'') <> '' ";
            strSql = strSql + " ORDER BY [User_NO] ASC";
            dt = cmd.GetDataTable("UserInfo", strSql);

            foreach (DataRow row in dt.Rows)
            {
                UserInfo userinfo = new UserInfo();
                userinfo.User_No = row["User_NO"].ToString();
                userinfo.User_Name = row["User_Name"].ToString();
                list.Add(userinfo);
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return list;
    }




}
