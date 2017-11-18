using com.data.structure;
using com.tools.db;
using com.tools.file;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.ServiceModel.Activation;

[AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Required)]
// 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码、服务和配置文件中的类名“Service”。
public class Service : IService
{
    System.DateTime currentTime = new System.DateTime();
    private string programname = "Service";
    private string errlog = "";

    public DataTable GetUserData(string input = "")
    {
        DataTable dt = new DataTable();
        string strSql = "";

        string[] data = input.Split(new Char[] { ',' });
        string user_enname = data[0].Trim();
        string user_no = data[1].Trim();

        try
        {
            strSql = "";
            strSql = strSql + " SELECT";
            strSql = strSql + " User_NO AS [User_NO],";
            strSql = strSql + " User_EnName AS [User_EnName],";
            strSql = strSql + " User_CnName AS [User_CnName],";
            strSql = strSql + " (SELECT Sys_ItemName FROM dbo.[SystemCode] WHERE [Sys_Code] = 'User_SectionID' AND [Sys_ItemValue] = A.User_SectionID) AS User_Section";
            strSql = strSql + " FROM dbo.[UserTable] AS A";
            strSql = strSql + " WHERE A.[USER_TYPE] = '2' AND ISNULL(A.[User_GroupID],'') <>''";

            if (!string.IsNullOrEmpty(user_enname))
            {
                strSql = strSql + " AND A.[User_EnName] LIKE '%" + user_enname.Trim() + "%'";
            }

            strSql = strSql + " ORDER BY A.[User_NO] ASC";

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("UserData", strSql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public DataTable GetCheckUserAccount(string systemcode,string account, string password)
    {
        DataTable dt = new DataTable();
        string sql = "";

        try
        {
            if (account.Trim() != "" && password.Trim() != "")
            {
                Info info = new Info();
                Command cmd = new Command(info);
                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();

                switch (systemcode)
                {
                    case "Reporting":
                        sql = "";
                        sql = sql + " SELECT ";
                        sql = sql + "   User_NO, ";
                        sql = sql + "   User_ID, ";
                        sql = sql + "   User_GroupID, ";
                        sql = sql + "   User_CivilizID, ";
                        sql = sql + "   User_SectionID ";
                        sql = sql + " FROM UserTable ";
                        sql = sql + " WHERE User_ID = ? And User_Pwd =  ? ";
                        
                        OleDbParameters.Add(new OleDbParameter("@User_ID", account));
                        OleDbParameters.Add(new OleDbParameter("@User_Pwd", password));

                        dt = cmd.GetDataTable("UserTable", sql, OleDbParameters);
                        break;

                    case "Jdboarweb":
                        //先判断 帐号 密码（JdbOadata） 对不对，再取 UserCenter 的 UserTable

                        sql = "";
                        sql = sql + " SELECT ";
                        sql = sql + "   User_NO, ";
                        sql = sql + "   User_ID, ";
                        sql = sql + "   User_GroupID, ";
                        sql = sql + "   User_CivilizID, ";
                        sql = sql + "   User_SectionID ";
                        sql = sql + " FROM [dbo].[View_UserTable] ";
                        sql = sql + " WHERE User_NO = ? And user_passwd = ? ";

                        OleDbParameters.Add(new OleDbParameter("@User_NO", account));
                        OleDbParameters.Add(new OleDbParameter("@user_passwd", password));

                        dt = cmd.GetDataTable("UserTable", sql, OleDbParameters);
                        break;
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

    public DataTable getUserArea(string userNo)
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            if (userNo.Trim() != "")
            {
                strSQL = strSQL + "SELECT * FROM [dbo].[View_User_Area_Branch] WHERE BR_User_NO = ? ";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@BR_User_NO", userNo.Trim()));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("UserMenuRule", strSQL, OleDbParameters);
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

    public DataTable getUserMenuRule(string groupID, string userNo)
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            if (groupID.Trim() != "" && userNo.Trim() != "")
            {
                strSQL = strSQL + " SELECT ";
                strSQL = strSQL + " GR_School_TypeID AS School_Type_ID,";
                strSQL = strSQL + " (SELECT Sys_ItemName FROM SystemCode WHERE Sys_Code='School_TypeID' AND Sys_ItemValue = A.GR_School_TypeID) AS School_Type, ";
                strSQL = strSQL + " B.Menu_GroupID, ";
                strSQL = strSQL + " (SELECT Sys_ItemName FROM SystemCode WHERE Sys_Code='Menu_GroupID' AND Sys_ItemValue = B.Menu_GroupID) AS Menu_Group,";
                strSQL = strSQL + " GR_Menu_Code AS Menu_Code, ";
                strSQL = strSQL + " B.Menu_Sort, ";
                strSQL = strSQL + " B.Menu_Desc_TW, ";
                strSQL = strSQL + " B.Menu_Desc_CN, ";
                strSQL = strSQL + " B.Menu_Path";
                strSQL = strSQL + " FROM [dbo].[Group_Rule] AS A LEFT JOIN [dbo].[Menu] AS B ON A.GR_School_TypeID = B.Menu_School_TypeID AND A.GR_Menu_Code = B.Menu_Code ";
                strSQL = strSQL + " WHERE GR_Group_ID = ? ";
                strSQL = strSQL + " UNION ";
                strSQL = strSQL + " SELECT ";
                strSQL = strSQL + " MR_School_TypeID AS School_Type_ID,";
                strSQL = strSQL + " (SELECT Sys_ItemName FROM SystemCode WHERE Sys_Code='School_TypeID' AND Sys_ItemValue = C.MR_School_TypeID) AS School_Type,";
                strSQL = strSQL + " D.Menu_GroupID,";
                strSQL = strSQL + " (SELECT Sys_ItemName FROM SystemCode WHERE Sys_Code='Menu_GroupID' AND Sys_ItemValue = D.Menu_GroupID) AS Menu_Group,";
                strSQL = strSQL + " MR_Menu_Code AS Menu_Code,";
                strSQL = strSQL + " D.Menu_Sort,";
                strSQL = strSQL + " D.Menu_Desc_TW,";
                strSQL = strSQL + " D.Menu_Desc_CN,";
                strSQL = strSQL + " D.Menu_Path";
                strSQL = strSQL + " FROM [dbo].[Menu_Rule] AS C LEFT JOIN [dbo].[Menu] AS D ON C.MR_School_TypeID = D.Menu_School_TypeID AND C.MR_Menu_Code = D.Menu_Code ";
                strSQL = strSQL + " WHERE MR_User_NO = ? ";
                strSQL = strSQL + " ORDER BY School_Type_ID,Menu_GroupID,Menu_Sort ASC";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@GR_Group_ID", groupID.Trim()));
                OleDbParameters.Add(new OleDbParameter("@MR_User_NO", userNo.Trim()));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("UserMenuRule", strSQL, OleDbParameters);
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

    public DataTable getMenu(string School_TypeID)
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            strSQL = strSQL + " SELECT DISTINCT";
            strSQL = strSQL + " Menu_SystemID, ";
            strSQL = strSQL + " Menu_GroupID, ";
            strSQL = strSQL + " (SELECT Sys_ItemName  FROM [dbo].[SystemCode] WHERE Sys_Code = 'Menu_GroupID' AND Sys_ItemValue = Menu.Menu_GroupID) AS Menu_Group, ";
            strSQL = strSQL + " Menu_Code, ";
            strSQL = strSQL + " Menu_Desc_TW, ";
            strSQL = strSQL + " Menu_Desc_CN, ";
            strSQL = strSQL + " Menu_Sort, ";
            strSQL = strSQL + " Menu_Path ";
            strSQL = strSQL + " From Menu ";
            if (School_TypeID.Trim() != "")
            {
                strSQL = strSQL + " WHERE Menu_School_TypeID = ? ";
            }
            strSQL = strSQL + " Order By Menu_SystemID,Menu_GroupID,Menu_Sort ASC ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Menu_School_TypeID", School_TypeID.Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Menu", strSQL, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public DataTable getMenuByGroup(string School_TypeID, string User_GroupID)
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            strSQL = strSQL + " SELECT DISTINCT";
            strSQL = strSQL + " (Case WHEN (SELECT GR_Menu_Code FROM [dbo].[Group_Rule] WHERE GR_School_TypeID = ? AND GR_Group_ID = ? AND GR_Menu_Code = A.Menu_Code) = A.Menu_Code  THEN 'True'";
            strSQL = strSQL + " Else";
            strSQL = strSQL + " 	'False'";
            strSQL = strSQL + " End) AS Checked,";
            strSQL = strSQL + " Menu_SystemID, ";
            strSQL = strSQL + " Menu_GroupID, ";
            strSQL = strSQL + " (SELECT Sys_ItemName  FROM [dbo].[SystemCode] WHERE Sys_Code = 'Menu_GroupID' AND Sys_ItemValue = Menu.Menu_GroupID) AS Menu_Group, ";
            strSQL = strSQL + " Menu_Code, ";
            strSQL = strSQL + " Menu_Desc_TW, ";
            strSQL = strSQL + " Menu_Desc_CN, ";
            strSQL = strSQL + " Menu_Sort, ";
            strSQL = strSQL + " Menu_Path ";
            strSQL = strSQL + " From Menu ";
            strSQL = strSQL + " WHERE 1=1 ";
            if (School_TypeID.Trim() != "")
            {
                strSQL = strSQL + " WHERE Menu_School_TypeID = ? ";
            }
            strSQL = strSQL + " Order By Menu_SystemID,Menu_GroupID,Menu_Sort ASC ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@GR_School_TypeID", School_TypeID.Trim()));
            OleDbParameters.Add(new OleDbParameter("@GR_Group_ID", User_GroupID.Trim()));
            OleDbParameters.Add(new OleDbParameter("@Menu_School_TypeID", School_TypeID.Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Menu", strSQL, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public DataTable getMenu_Rule(string input)
    {
        DataTable dt = new DataTable();
        string strSQL = "";
        string[] data = input.Split(new Char[] { ',' });
        string schooltypeid = data[0].Trim();
        string usergroup = data[1].Trim();
        string userno = data[2].Trim();
        string menupath = data[3].Trim();

        try
        {
            strSQL = strSQL + " SELECT DISTINCT ";
            // Vesper 补上群组的 新增 修改 删除 列印 退回 等权限功能 20141013
            strSQL = strSQL + "  (SELECT User_Type FROM UserTable WHERE User_NO = A.MR_User_NO) AS User_Type, ";
            //
            strSQL = strSQL + "  A.[MR_Menu_Code],A.[MR_SELECT],A.[MR_Insert],A.[MR_Delete],[MR_Update],A.[MR_Print],A.[MR_Return] ";
            strSQL = strSQL + " FROM  [dbo].[Menu_Rule] AS A LEFT JOIN [dbo].[Menu] AS B ON A.[MR_Menu_Code] = B.[Menu_Code] ";
            strSQL = strSQL + " WHERE 1=1 ";

            if (!string.IsNullOrEmpty(schooltypeid))
            {
                strSQL = strSQL + " AND  A.MR_School_TypeID ='" + schooltypeid + "' ";
            }

            if (!string.IsNullOrEmpty(userno))
            {
                // Vesper 补上群组的 新增 修改 删除 列印 退回 等权限功能 20141013
                strSQL = strSQL + " AND (A.MR_User_NO = '" + usergroup + "' OR A.MR_User_NO = '" + userno + "') ";
                //
            }

            if (!string.IsNullOrEmpty(menupath))
            {
                strSQL = strSQL + " AND B.Menu_Path like '%" + menupath + "' ";
            }

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Menu_Rule", strSQL);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }

        if (dt.Rows.Count >= 2 && !string.IsNullOrEmpty(menupath))
        {
            //var query = dt.Select("User_Type = '2'");
            //return query.CopyToDataTable();

            DataTable newdatatabele = new DataTable();
            newdatatabele = dt.Clone();
            foreach (DataRow dr in dt.Rows)
            {
                if (dr["User_Type"].ToString() == "2")
                {
                    newdatatabele.ImportRow(dr);
                }
            }
            return newdatatabele;
        }
        else
        {
            return dt;
        }
    }

    public DataTable getArea_Branch()
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            strSQL = strSQL + " SELECT ID,";
            strSQL = strSQL + " Left(SysCode.Sys_ItemValue,2) AS Area_ID,";
            strSQL = strSQL + " (SELECT Sys_ItemName  FROM [dbo].[SystemCode] WHERE Sys_Code = 'Area' AND Sys_ItemValue = Left(SysCode.Sys_ItemValue,2)) AS Area,";
            strSQL = strSQL + " (SUBSTRING(SysCode.Sys_ItemValue,3,1)) AS School_TypeID,";
            strSQL = strSQL + " (SELECT Sys_ItemName  FROM [dbo].[SystemCode] WHERE Sys_Code = 'School_TypeID' AND Sys_ItemValue = SUBSTRING(SysCode.Sys_ItemValue,3,1)) AS School_Type,";
            strSQL = strSQL + " Sys_ItemValue AS Branch_ID,";
            strSQL = strSQL + " Sys_ItemName AS Branch_Desc";
            strSQL = strSQL + " FROM [dbo].[SystemCode] AS SysCode  WHERE Sys_Code='Branch'";
            strSQL = strSQL + " ORDER BY Area_ID,School_TypeID,Branch_ID ASC ";

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Area_Branch", strSQL);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public DataTable getSystem_Code(string Sys_Code = "")
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            strSQL = strSQL + " SELECT ";
            strSQL = strSQL + "  [ID], ";
            strSQL = strSQL + "  Sys_Code,";
            strSQL = strSQL + "  Sys_Name,";
            strSQL = strSQL + "  Sys_ItemValue,";
            strSQL = strSQL + "  Sys_ItemName";
            strSQL = strSQL + "  FROM SystemCode";
            if (Sys_Code.Trim() != "")
            {
                strSQL = strSQL + "  WHERE Sys_Code =  ? ";
            }
            strSQL = strSQL + "  ORDER BY Sys_Code,Sys_ItemValue ASC";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Sys_Code", Sys_Code.Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("System_Code", strSQL, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public DataTable getUserBranchs(string input)
    {
        DataTable dt = new DataTable();
        string sql = "";
        string[] Branchs = input.Split(new Char[] { ',' });

        try
        {
            sql = sql + " SELECT ID,";
            sql = sql + "  Sys_ItemValue,";
            sql = sql + "  Sys_ItemName";
            sql = sql + "  FROM SystemCode";
            sql = sql + "  WHERE Sys_Code = 'Branch' ";

            if (!string.IsNullOrEmpty(input))
            {
                sql = sql + " AND Sys_ItemValue IN (";
                for (int i = 0; i <= Branchs.Length - 1; i++)
                {
                    sql = sql + "'" + Branchs[i].ToString() + "',";
                }
                sql = sql.Substring(0, sql.Length - 1);
                sql = sql + ")";
            }
            sql = sql + "  ORDER BY Sys_ItemValue ASC";

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Area_Branch", sql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public DataTable getUserBranchsBySchoolType(string input)
    {
        DataTable dt = new DataTable();
        string sql = "";

        try
        {
            sql = sql + " SELECT ID,";
            sql = sql + "  Sys_ItemValue,";
            sql = sql + "  Sys_ItemName";
            sql = sql + "  FROM SystemCode";
            sql = sql + "  WHERE Sys_Code = 'Branch' ";

            if (!string.IsNullOrEmpty(input))
            {
                sql = sql + " AND Sys_ItemValue Like '%__" + input + "__%'";
            }
            sql = sql + "  ORDER BY Sys_ItemValue ASC";

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Area_Branch", sql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public DataTable getGroupRule(string SchoolType, string GroupID)
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            strSQL = strSQL + " SELECT GR_Menu_Code ";
            strSQL = strSQL + " FROM [dbo].[Group_Rule]";
            strSQL = strSQL + " WHERE GR_School_TypeID = ? AND GR_Group_ID = ? ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@GR_School_TypeID", SchoolType.Trim()));
            OleDbParameters.Add(new OleDbParameter("@GR_Group_ID", GroupID.Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Area_Branch", strSQL, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public string getUserInfo(string User_NO = "", string User_EnName = "")
    {
        string result = "";
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            strSQL = strSQL + " SELECT User_NO,User_CnName ";
            strSQL = strSQL + " FROM [dbo].[UserTable]";
            strSQL = strSQL + " WHERE 1=1 ";

            if (String.IsNullOrEmpty(User_NO.Trim()) != true)
            {
                strSQL = strSQL + "And User_No= ? ";
            }

            if (String.IsNullOrEmpty(User_EnName.Trim()) != true)
            {
                strSQL = strSQL + "And User_EnName= ? ";
            }

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@User_No", User_NO.Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_EnName", User_EnName.Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("UserTable", strSQL, OleDbParameters);

            if (dt.Rows.Count == 1)
            {
                result = dt.Rows[0]["User_NO"].ToString().Trim() + "_" + dt.Rows[0]["User_CnName"].ToString().Trim();
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

    public int setGroupRule(string SchoolType, string GroupID, string[] Menu_Code)
    {
        int result = 0;
        string strSQL = "";

        try
        {
            strSQL = strSQL + " BEGIN TRAN ";
            strSQL = strSQL + " DELETE FROM [Group_Rule] WHERE GR_School_TypeID = '" + SchoolType.Trim() + "' AND GR_Group_ID='" + GroupID.Trim() + "' ";

            for (int i = 0; i <= Menu_Code.Length - 1; i++)
            {
                strSQL = strSQL + " Insert Group_Rule(GR_School_TypeID,GR_Group_ID,GR_Menu_Code) Values('" + SchoolType.Trim() + "','" + GroupID.Trim() + "','" + Menu_Code[i].ToString().Trim() + "') ";
            }
            strSQL = strSQL + " COMMIT TRAN ";

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

    public DataTable getBranchRule(string User_No)
    {
        DataTable dt = new DataTable();
        string strSQL = "";

        try
        {
            strSQL = strSQL + " SELECT BR_Branch_ID,BR_User_NO ";
            strSQL = strSQL + " FROM [dbo].[Branch_Rule] ";
            strSQL = strSQL + " WHERE BR_User_NO = ? ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@BR_User_NO", User_No.Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Branch_Rule", strSQL, OleDbParameters);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public int setBranchRule(string userno, string branchs, string type)
    {
        int result = 0;
        string area = "";
        string school_type = "";
        string branch = "";
        string[] Branch_ID = branchs.Split(new Char[] { ',' });

        try
        {
            for (int i = 0; i <= Branch_ID.Length - 1; i++)
            {
                if (!string.IsNullOrEmpty(Branch_ID[i].ToString().Trim()))
                {
                    branch = Branch_ID[i].ToString().Trim();
                    area = branch.Substring(0, 2);
                    school_type = branch.Substring(2, 1);

                    if (type == "INSERT")
                    {
                        InsertArea(area.Trim(), school_type.Trim(), userno.Trim());
                        result = result + InsertBranch(branch.Trim(), userno.Trim());
                    }

                    if (type == "DELETE")
                    {
                        DeleteArea(area.Trim(), school_type.Trim(), userno.Trim());
                        result = result + DeleteBranch(branch.Trim(), userno.Trim());
                    }
                }
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

    private int InsertBranch(string branch, string userno)
    {
        int result = 0;
        string sql = "";
        DataTable dt = new DataTable();

        try
        {
            if (!string.IsNullOrEmpty(branch.Trim()) && !string.IsNullOrEmpty(userno.Trim()))
            {
                sql = sql + " SELECT COUNT(*) FROM [dbo].[Branch_Rule] WHERE BR_Branch_ID = ? AND BR_User_NO = ? ";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@BR_Branch_ID", branch.Trim()));
                OleDbParameters.Add(new OleDbParameter("@BR_User_NO", userno.Trim()));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("Branch_Rule", sql, OleDbParameters);

                if ((int)dt.Rows[0][0] <= 0)
                {
                    sql = "";
                    sql = sql + "INSERT INTO [Branch_Rule](BR_Branch_ID,BR_User_NO) VALUES(?,?) ";
                    result = cmd.ExecuteSQL(sql, OleDbParameters);
                }
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

    private int InsertArea(string area, string schooltype, string userno)
    {
        string strSQL = "";
        int result = 0;
        DataTable dt = new DataTable();

        try
        {
            if (!string.IsNullOrEmpty(area.Trim()) && !string.IsNullOrEmpty(schooltype.Trim()) && !string.IsNullOrEmpty(userno.Trim()))
            {
                strSQL = strSQL + "SELECT Count(*) FROM [dbo].[Area_Rule] WHERE Area_ID= ? AND Area_School_TypeID = ? AND Area_User_NO = ? ";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@Area_ID", area.Trim()));
                OleDbParameters.Add(new OleDbParameter("@Area_School_TypeID", schooltype.Trim()));
                OleDbParameters.Add(new OleDbParameter("@Area_User_NO", userno.Trim()));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("Branch_Rule", strSQL, OleDbParameters);                

                if ((int)dt.Rows[0][0] <= 0)
                {
                    strSQL = "";
                    strSQL = strSQL + " INSERT INTO [dbo].[Area_Rule](Area_ID,Area_School_TypeID,Area_User_NO) VALUES(?,?,?) ";
                    result = cmd.ExecuteSQL(strSQL,OleDbParameters);
                }
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

    private int DeleteArea(string area, string schooltype, string userno)
    {
        string strSQL = "";
        int result = 0;
        DataTable dt = new DataTable();

        try
        {
            if (!string.IsNullOrEmpty(area.Trim()) && !string.IsNullOrEmpty(schooltype.Trim()) && !string.IsNullOrEmpty(userno.Trim()))
            {
                strSQL = strSQL + "SELECT Count(*) FROM [dbo].[Area_Rule] WHERE Area_ID= ? AND Area_School_TypeID = ? AND Area_User_NO = ? ";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@Area_ID", area.Trim()));
                OleDbParameters.Add(new OleDbParameter("@Area_School_TypeID", schooltype.Trim()));
                OleDbParameters.Add(new OleDbParameter("@Area_User_NO", userno.Trim()));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("Area_Rule", strSQL, OleDbParameters);   

                if ((int)dt.Rows[0][0] == 1)
                {
                    strSQL = "";
                    strSQL = strSQL + " DELETE FORM [dbo].[Area_Rule] WHERE Area_ID= ? AND Area_School_TypeID = ? AND Area_User_NO =  ? ";
                    result = cmd.ExecuteSQL(strSQL,OleDbParameters);
                }
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

    private int DeleteBranch(string branch, string userno)
    {
        int result = 0;
        string sql = "";
        DataTable dt = new DataTable();

        try
        {
            if (!string.IsNullOrEmpty(branch.Trim()) && !string.IsNullOrEmpty(userno.Trim()))
            {
                sql = sql + " SELECT COUNT(*) FROM [dbo].[Branch_Rule] WHERE BR_Branch_ID = ? AND BR_User_NO = ? ";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@BR_Branch_ID", branch.Trim()));
                OleDbParameters.Add(new OleDbParameter("@BR_User_NO", userno.Trim()));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("Branch_Rule", sql, OleDbParameters);   

                if ((int)dt.Rows[0][0] == 1)
                {
                    sql = "";
                    sql = sql + "DELETE FROM [Branch_Rule] WHERE BR_Branch_ID = ? AND BR_User_NO = ? ";
                    result = cmd.ExecuteSQL(sql,OleDbParameters);
                }
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

    public int InsertSysCode(string input)
    {
        int result = 0;
        string sql = "";
        string[] Sys_Code = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + "INSERT INTO [dbo].[SystemCode](Sys_Code,Sys_Name,Sys_ItemValue,Sys_ItemName) VALUES(?,?,?,?) ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Sys_Code", Sys_Code[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Sys_Name", Sys_Code[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Sys_ItemValue", Sys_Code[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Sys_ItemName", Sys_Code[3].Trim()));

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

    public int UpdateSysCode(string input)
    {
        int result = 0;
        string sql = "";
        string[] Sys_Code = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " UPDATE [dbo].[SystemCode] SET ";
            sql = sql + "   Sys_Code = ?,";
            sql = sql + "   Sys_Name = ?,";
            sql = sql + "   Sys_ItemValue = ?,";
            sql = sql + "   Sys_ItemName = ?";
            sql = sql + " WHERE ID= ? ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Sys_Code", Sys_Code[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Sys_Name", Sys_Code[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Sys_ItemValue", Sys_Code[3].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Sys_ItemName", Sys_Code[4].Trim()));
            OleDbParameters.Add(new OleDbParameter("@ID", Sys_Code[0].Trim()));

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

    public int DeleteSysCode(string input)
    {
        int result = 0;
        string sql = "";

        try
        {
            if (!string.IsNullOrEmpty(input.Trim()))
            {
                sql = "";
                sql = sql + " DELETE FROM  [dbo].[SystemCode] WHERE ID='" + input.Trim() + "' ";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@ID", input.Trim()));

                Info info = new Info();
                Command cmd = new Command(info);
                result = cmd.ExecuteSQL(sql,OleDbParameters);
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

    public DataTable getUserTable(string input = "")
    {
        DataTable dt = new DataTable();
        string sql = "";
        string usertype = null;
        string userno = null;
        string school_type = null;

        if (!string.IsNullOrEmpty(input))
        {
            string[] data = input.Split(new Char[] { ',' });
            usertype = data[0].Trim();
            userno = data[1].Trim();
            school_type = data[2].Trim();
        }

        try
        {
            sql = sql + " SELECT ";
            sql = sql + " [ID],";
            sql = sql + " [User_NO],";
            sql = sql + " [User_ID],";
            sql = sql + " [User_Pwd],";
            sql = sql + " [User_GroupID],";
            sql = sql + " (SELECT Sys_ItemName FROM dbo.SystemCode WHERE Sys_Code='User_GroupID' AND Sys_ItemValue = [UserTable].[User_GroupID]) AS [User_Group],";
            sql = sql + " [User_CnName],";
            sql = sql + " [User_EnName],";
            sql = sql + " Convert(varchar(10),[User_Birthday],111) AS [User_Birthday],";
            sql = sql + " [User_GroupID],";
            sql = sql + " [User_SectionID],";
            sql = sql + " (SELECT Sys_ItemName FROM dbo.SystemCode WHERE Sys_Code='User_SectionID' AND Sys_ItemValue = [UserTable].[User_SectionID]) AS [User_Section] ";
            sql = sql + " FROM [UserTable]";
            sql = sql + " WHERE 1=1  AND ISNULL(User_ID,'') <> '' ";
            //VESPER 暂时将没使用的使用者过滤掉 20141205
            sql = sql + " AND  ISNULL(User_GroupID,'') <>'' ";

            if (!string.IsNullOrEmpty(usertype))
            {
                sql = sql + " AND [User_Type] = '" + usertype.Trim() + "' ";
            }

            if (!string.IsNullOrEmpty(userno))
            {
                sql = sql + " AND [User_NO] = '" + userno.Trim() + "' ";
            }

            if (!string.IsNullOrEmpty(school_type))
            {
                //VESPER 因数剧量太大 所以以公司类别来取 20141114
                //以员工编码来判断 公司类别
                sql = sql + " AND [User_NO] like '%__" + school_type.Trim() + "_______%' ";
            }

            sql = sql + " Order By User_No,User_ID ASC";
            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("UserTable", sql);
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);
        }
        return dt;
    }

    public int InsertUserTable(string input)
    {
        int result = 0;
        string sql = "";
        string[] UserData = input.Split(new Char[] { ',' });

        try
        {
            sql = "";
            sql = sql + " INSERT INTO UserTable(User_Type,User_NO,User_ID,User_Pwd,User_GroupID,User_CnName,User_EnName,User_Birthday,User_SectionID,Ins_UserID,Ins_DateTime)";
            sql = sql + " VALUES(?,?,?,?,?,?,?,?,?,?,?)";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@User_Type", UserData[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_NO", UserData[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_ID", UserData[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_Pwd", UserData[3].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_GroupID", UserData[4].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_CnName", UserData[5].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_EnName", UserData[6].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_Birthday", UserData[7].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_SectionID", UserData[8].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Ins_UserID", UserData[9].Trim()));
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

    public int UpdateUserTable(string input)
    {
        int result = 0;
        string sql = "";
        string[] UserData = input.Split(new Char[] { ',' });

        try
        {
            sql = sql + " UPDATE UserTable Set";
            sql = sql + " User_ID = ?,";
            sql = sql + " User_Pwd = ?,";
            sql = sql + " User_GroupID = ?,";
            sql = sql + " User_CnName = ?,";
            sql = sql + " User_EnName = ?,";
            sql = sql + " User_Birthday = ?,";
            sql = sql + " User_SectionID = ?,";
            sql = sql + " Mdy_UserID = ?,";
            sql = sql + " Mdy_DateTime = ?";
            sql = sql + " WHERE  ID = ?";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@User_ID", UserData[0].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_Pwd", UserData[1].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_GroupID", UserData[2].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_CnName", UserData[3].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_EnName", UserData[4].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_Birthday", UserData[5].Trim()));
            OleDbParameters.Add(new OleDbParameter("@User_SectionID", UserData[6].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Mdy_UserID", UserData[8].Trim()));
            OleDbParameters.Add(new OleDbParameter("@Mdy_DateTime", System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss")));
            OleDbParameters.Add(new OleDbParameter("@ID", int.Parse(UserData[7].Trim())));

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

    public int DeleteUserTable(string userno)
    {
        int result = 0;
        string sql = "";

        try
        {
            sql = sql + " DELETE FROM  [dbo].[UserTable] WHERE ID= ? ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@ID", userno.Trim()));

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


    public string getUserBranchsByString(string schooltypeid, string userid)
    {
        string result = "";
        DataTable dt = new DataTable();
        dt = getUserArea(userid);

        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                // 只取学校类别相同的 学校 或 园
                if (dr["Area_School_TypeID"].ToString() == schooltypeid.Trim())
                {
                    result = result + dr["BR_Branch_ID"].ToString() + ",";
                }
            }

            if (result.IndexOf(",") > 0)
            {
                //去掉最后一个","号
                result = result.Substring(0, result.Length - 1);
            }
        }
        return result;
    }

    public int setMenu_Rule(string input)
    {
        int result = 0;
        string sql = "";
        string[] Data = input.Split(new Char[] { '|' });

        string schooltype = Data[0].ToString().Trim();
        string userno = Data[1].ToString().Trim();
        string insuserno = Data[2].ToString().Trim();
        string jsonstring = Data[3].ToString().Trim();

        try
        {
            sql = sql + " DELETE FROM  [dbo].[Menu_Rule] WHERE MR_School_TypeID='" + schooltype + "' AND MR_User_NO = '" + userno + "' " + "\r\n";

            // Vesper jsonstring 不为空集合时才做 20141013
            if (jsonstring != "[}")
            {
                List<Menu_Rule> o = new List<Menu_Rule>();
                o = com.tools.json.josnconvert.ConvertJsonStringToList<Menu_Rule>(jsonstring);

                foreach (Menu_Rule m in o)
                {
                    sql = sql + "INSERT INTO [dbo].[Menu_Rule](MR_School_TypeID,MR_Menu_Code,MR_User_NO,MR_SELECT,MR_Insert,MR_Delete,MR_Update,MR_Print,MR_Return,Ins_UserID,Ins_DateTime)";
                    sql = sql + " VALUES('" + schooltype + "','" + m.MR_Menu_Code.Trim() + "','" + userno + "','" + m.MR_SELECT.Trim() + "','" + m.MR_Insert + "','" + m.MR_Delete.Trim() + "','" + m.MR_Update.Trim() + "','" + m.MR_Print.Trim() + "','" + m.MR_Return.Trim() + "','" + insuserno + "','" + System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + "' )" + "\r\n";
                }
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

    /// <summary>
    /// 提供给 Jdboarweb 外部系统 开起 报表表系统 部份功能用
    /// </summary>
    /// <param name="input">1|AndyHuang|KidCastle1986|Report</param>
    /// <returns>
    /// ""：未传参数
    /// "1"：登入失败
    /// "2"：没有浏览权限
    /// "errlog:"：相关错误信息
    /// "http://"：开起功能的网址
    /// </returns>

    public string getUrlFromJdboarweb(string input)
    {
        string result = "";
        string[] Data = input.Split(new Char[] { '|' });
        string branchs = "";

        try
        {
            if (Data.Length > 0)
            {
                string school_typeid = Data[0].ToString().Trim(); //学校类别 1:公司 2：园 3：证照校 4：校
                string user_no = Data[1].ToString().Trim();//帐号：
                string user_pwd = Data[2].ToString().Trim();//密码：
                string category = Data[3].ToString().Trim();//回传类别：预留以后其它功能可能会用 目前只用于报表

                DataTable dt = GetCheckUserAccount("Jdboarweb", user_no, user_pwd);

                if (dt != null)
                {

                    // VESPER 如果学校类别 等于 公司 或 证照校，则去取该使用者能登的学校 20141205
                    //if (school_typeid == "1" || school_typeid == "3")
                    //{
                    //    branchs = getUserBranchsByString("4", dt.Rows[0]["User_NO"].ToString().Trim());
                    //}
                    //else
                    //{
                    //    branchs = getUserBranchsByString(school_typeid, dt.Rows[0]["User_NO"].ToString().Trim());
                    //}

                    //Vesper 读取 Jdbo 的学校权限 20150427
                    //branchs = getBranchFromJdbo(user_no);
                    // 将 branchs 加密 20150511
                    branchs = OperatorFile.Encrypt(getBranchFromJdbo(user_no));

                    if (!string.IsNullOrEmpty(branchs))
                    {
                        result = "http://localhost:1643/jdboareport/jReport00001.aspx?Branchs=" + branchs;
                    }
                    else
                    {
                        result = "2"; //没有浏览权限
                    }
                }
                else
                {
                    result = "1"; //登入失败
                }
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);

            result = getErrorCode(ex.Message.ToString());
        }
        return result;
    }

    private string getBranchFromJdbo(string user_no) 
    {
        string result = "";
        string sql = "";
        DataTable dt = new DataTable();
        int rowcount = 0;
        
        try 
        {
            sql = sql + " SELECT group_pid FROM [dbo].[View_User_Branch_Jdbo]  WHERE user_pid = ? ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@user_pid", user_no.Trim()));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("View_User_Branch_Jdbo", sql, OleDbParameters);            

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    result =result + dr["group_pid"].ToString().Trim() + ((dt.Rows.Count != rowcount)?",":"");
                    rowcount = rowcount + 1;
                }
            }
        }
        catch (Exception ex)
        {
            errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
            OperatorFile OF = new OperatorFile();
            OF.WriteAppend(errlog);

            result = getErrorCode(ex.Message.ToString());
        }
        return result;
    }


    private string getErrorCode(string message) 
    {
        string result = "";

        switch (message) 
        {
            case "There is no row at position 0.":
                result = "3"; // 在报表平台没有帐号
                break;
        }
        return result;
    }


    public int InsertAttachment(AttachmentDataModel attachment) 
    {
        int result = 0;
        string sql = "";

        try 
        {
            sql = sql + " INSERT INTO Attachment([Category],[FromID],[Path],[UploadDateTime])";
            sql = sql + " VALUES(?,?,?,?)";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Category", attachment.Category));
            OleDbParameters.Add(new OleDbParameter("@FromID", int.Parse(attachment.FromID)));
            OleDbParameters.Add(new OleDbParameter("@Path", attachment.Path));
            OleDbParameters.Add(new OleDbParameter("@UploadDateTime", System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss")));

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

    public string getSystemCode_ItemName(string input) 
    {
        string result = "";
        string sql = "";
        DataTable dt = new DataTable();

        string[] Data = input.Split(new Char[] { '|' });
        string systemcode = Data[0].ToString().Trim();//帐号：
        string itemvalue = Data[1].ToString().Trim();//密码：

        try
        {
            sql = sql + " SELECT ";
            sql = sql + "  Sys_ItemName";
            sql = sql + "  FROM SystemCode";
            if (!string.IsNullOrEmpty(systemcode) && !string.IsNullOrEmpty(itemvalue))
            {
                sql = sql + "  WHERE Sys_Code = '" + systemcode + "' And Sys_ItemValue ='" + itemvalue+"' ";
            }

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Sys_Code", systemcode));
            OleDbParameters.Add(new OleDbParameter("@Sys_ItemValue", itemvalue));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("System_Code", sql, OleDbParameters);

            if (dt.Rows.Count > 0) 
            {
                result = dt.Rows[0]["Sys_ItemName"].ToString().Trim();
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

    public int DeleteAttachment(string Category, string fromid)
    {
        int result = 0;
        string sql = "";

        try
        {
            sql = "";
            sql = sql + " DELETE FROM  [dbo].[Attachment] WHERE Category= ? AND  FromID = ? ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Category", Category));
            OleDbParameters.Add(new OleDbParameter("@FromID", fromid));

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

    public bool IsAttachment(string Category, string fromid) 
    {
        bool result = false;
        string sql = "";
        DataTable dt = new DataTable();

        try
        {
            sql = sql + " SELECT COUNT(ID) AS CNT";
            sql = sql + " FROM Attachment";
            sql = sql + " WHERE Category= ? AND  FromID = ? ";

            List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
            OleDbParameters.Add(new OleDbParameter("@Category", Category));
            OleDbParameters.Add(new OleDbParameter("@FromID", fromid));

            Info info = new Info();
            Command cmd = new Command(info);
            dt = cmd.GetDataTable("Attachment", sql, OleDbParameters);
            result = (int.Parse(dt.Rows[0]["CNT"].ToString()) == 0) ? false : true;
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
