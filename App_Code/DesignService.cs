using com.data.structure;
using com.tools.db;
using com.tools.file;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.ServiceModel.Activation;

namespace service.designservice
{
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]

    // 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码、svc 和配置文件中的类名“DesignService”。

    public class DesignService : IDesignService
    {
        private const string programname = "DesignService";
        private string errlog = "";

        public DataTable GetDesignList()
        {
            DataTable dt = new DataTable();
            string strSql = "";

            try
            {
                strSql = "";
                strSql = strSql + " SELECT";
                strSql = strSql + "     [ID],";
                strSql = strSql + " 	[Subject],";
                strSql = strSql + "     [Applicant],";
                strSql = strSql + "     (SELECT User_CnName + '(' + User_EnName +')' From [dbo].[UserTable] WHERE User_NO = [dbo].[DesignRequest].Applicant) AS Applicant_Name,";
                strSql = strSql + "     [Department],";
                strSql = strSql + "     (SELECT Sys_ItemName FROM [dbo].[SystemCode] WHERE Sys_Code = 'User_SectionID' AND Sys_ItemValue = [dbo].[DesignRequest].Department) AS Department_Name,";
                strSql = strSql + "     [FinishDate],";
                strSql = strSql + "     [Datelimite]";
                strSql = strSql + " FROM [dbo].[DesignRequest]";
                strSql = strSql + " WHERE 1=1 ";
                strSql = strSql + " ORDER BY [ID] ASC";

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("DesignRequest", strSql);
            }
            catch (Exception ex)
            {
                errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
                OperatorFile OF = new OperatorFile();
                OF.WriteAppend(errlog);
            }
            return dt;
        }

        public DataTable GetDesignDtlList(int mainid)
        {
            DataTable dt = new DataTable();
            string strSql = "";

            try
            {
                strSql = "";
                strSql = strSql + " SELECT";
                strSql = strSql + " [ID],";
                strSql = strSql + " [MainID] AS MainID,";
                strSql = strSql + " [Type],";
                strSql = strSql + " [Count],";
                strSql = strSql + " [Layout-High],";
                strSql = strSql + " [Layout-Width],";
                strSql = strSql + " [Layout-Unit],";
                strSql = strSql + " [Layout-Type],";
                strSql = strSql + " [Layout-Page],";
                strSql = strSql + " [FrontCover],";
                strSql = strSql + " [FrontCoverType],";
                strSql = strSql + " [Content],";
                strSql = strSql + " [ContentType],";
                strSql = strSql + " [FrontCover-FrontColor],";
                strSql = strSql + " [FrontCover-ObverseColor],";
                strSql = strSql + " [Content-FrontColor],";
                strSql = strSql + " [Content-ObverseColor],";
                strSql = strSql + " [SpacialColor],";
                strSql = strSql + " [SpacialColorCount],";
                strSql = strSql + " [Request],";
                strSql = strSql + " [AttachmentMemo],";
                strSql = strSql + " [AttachmentProvideDate],";
                strSql = strSql + " [CounterSing],";
                strSql = strSql + " [CounterSingType],";
                strSql = strSql + " [CounterSingMemo],";
                strSql = strSql + " [Status],";
                strSql = strSql + " [InitalDate],";
                strSql = strSql + " [InitalWorkTime],";
                strSql = strSql + " [InitalMemo],";
                strSql = strSql + " [FirstDate],";
                strSql = strSql + " [FirstWorkTime],";
                strSql = strSql + " [FirstMemo],";
                strSql = strSql + " [SecondDate],";
                strSql = strSql + " [SecondWorkTime],";
                strSql = strSql + " [SecondMemo],";
                strSql = strSql + " [ThirdDate],";
                strSql = strSql + " [ThirdWorkTime],";
                strSql = strSql + " [ThirdMemo],";
                strSql = strSql + " [FinishDate],";
                strSql = strSql + " [Acceptance],";
                strSql = strSql + " [Designer]";
                strSql = strSql + " FROM [dbo].[DesignRequestDetail]";
                strSql = strSql + " WHERE [MainID] = ? ";
                strSql = strSql + " ORDER BY ID ASC";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@MainID", mainid));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("DesignRequestDetail", strSql, OleDbParameters);
            }
            catch (Exception ex)
            {
                errlog = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + " " + programname + " " + ex.Message.ToString();
                OperatorFile OF = new OperatorFile();
                OF.WriteAppend(errlog);
            }
            return dt;
        }

        public int DesignInsert(DesignDataModel design)
        {
            int result = 0;
            string sql = "";
            try 
            {
                //sql = sql + " INSERT INTO DesignRequest([Subject],[Applicant],[Department],[FinishDate],[Datelimite])";
                //sql = sql + " VALUES(?,?,?,?,?)";


                sql = sql + " INSERT INTO DesignRequest([Subject],[Applicant],[Department],[Datelimite])";
                sql = sql + " VALUES(?,?,?,?)";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@Subject", design.Subject));
                OleDbParameters.Add(new OleDbParameter("@Applicant", design.Applicant));
                OleDbParameters.Add(new OleDbParameter("@Department", design.Department));
                //OleDbParameters.Add(new OleDbParameter("@FinishDate", design.FinishDate));
                OleDbParameters.Add(new OleDbParameter("@Datelimite", design.Datelimite));

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

        public int DesignUpdate(DesignDataModel design)
        {
            int result = 0;
            string sql = "";
            try
            {
                sql = sql + " UPDATE dbo.[DesignRequest] Set";
                sql = sql + "   [Subject] = ?,";
                sql = sql + "   [Applicant] = ?,";
                sql = sql + "   [Department] = ?,";
                sql = sql + "   [FinishDate] = ?,";
                sql = sql + "   [Datelimite] = ?";
                sql = sql + " WHERE  ID = ?";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@Subject", design.Subject));
                OleDbParameters.Add(new OleDbParameter("@Applicant", design.Applicant));
                OleDbParameters.Add(new OleDbParameter("@Department", design.Department));
                OleDbParameters.Add(new OleDbParameter("@FinishDate", design.FinishDate));
                OleDbParameters.Add(new OleDbParameter("@Datelimite", design.Datelimite));
                OleDbParameters.Add(new OleDbParameter("@ID", design.ID));

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

        public int DesignDelete(int id)
        {
            int result = 0;
            string sql = "";

            try 
            {
                sql = "";
                sql = sql + " DELETE FROM DesignRequest WHERE ID =?";
                sql = sql + " DELETE FROM DesignRequestDetail WHERE MainID = ?";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@ID", id));
                OleDbParameters.Add(new OleDbParameter("@MainID", id));

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

        public int DesignDtlInsert(DesignDetailDataModel designDtl)
        {
            int result = 0;
            string sql = "";

            try 
            {
                sql = sql + " INSERT INTO [dbo].[DesignRequestDetail](";
                sql = sql + " [MainID],";
                sql = sql + " [Type],";
                sql = sql + " [Count],";
                sql = sql + " [Layout-High],";
                sql = sql + " [Layout-Width],";
                sql = sql + " [Layout-Unit],";
                sql = sql + " [Layout-Type],";
                sql = sql + " [Layout-Page],";
                sql = sql + " [FrontCover],";
                sql = sql + " [FrontCoverType],";
                sql = sql + " [Content],";
                sql = sql + " [ContentType],";
                sql = sql + " [FrontCover-FrontColor],";
                sql = sql + " [FrontCover-ObverseColor],";
                sql = sql + " [Content-FrontColor],";
                sql = sql + " [Content-ObverseColor],";
                sql = sql + " [SpacialColor],";
                sql = sql + " [SpacialColorCount],";
                sql = sql + " [Request],";
                sql = sql + " [AttachmentMemo],";
                sql = sql + " [AttachmentProvideDate],";
                sql = sql + " [CounterSing],";
                sql = sql + " [CounterSingType],";
                sql = sql + " [CounterSingMemo],";
                sql = sql + " [Status],";
                sql = sql + " [InitalDate],";
                sql = sql + " [InitalWorkTime],";
                sql = sql + " [InitalMemo],";
                sql = sql + " [FirstDate],";
                sql = sql + " [FirstWorkTime],";
                sql = sql + " [FirstMemo],";
                sql = sql + " [SecondDate],";
                sql = sql + " [SecondWorkTime],";
                sql = sql + " [SecondMemo],";
                sql = sql + " [ThirdDate],";
                sql = sql + " [ThirdWorkTime],";
                sql = sql + " [ThirdMemo],";
                sql = sql + " [FinishDate],";
                sql = sql + " [Acceptance],";
                sql = sql + " [Designer])";
                sql = sql + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@MainID", designDtl.MainID));
                OleDbParameters.Add(new OleDbParameter("@Type", designDtl.Type));
                OleDbParameters.Add(new OleDbParameter("@Count", designDtl.Count));
                OleDbParameters.Add(new OleDbParameter("@Layout-High", designDtl.Layout_High));
                OleDbParameters.Add(new OleDbParameter("@Layout-Width", designDtl.Layout_Width));
                OleDbParameters.Add(new OleDbParameter("@Layout-Unit", designDtl.Layout_Unit));
                OleDbParameters.Add(new OleDbParameter("@Layout-Type", designDtl.Layout_Type));
                OleDbParameters.Add(new OleDbParameter("@Layout-Page", designDtl.Layout_Page));
                OleDbParameters.Add(new OleDbParameter("@FrontCover", designDtl.FrontCover));
                OleDbParameters.Add(new OleDbParameter("@FrontCoverType", designDtl.FrontCoverType));
                OleDbParameters.Add(new OleDbParameter("@ContentType", designDtl.ContentType));
                OleDbParameters.Add(new OleDbParameter("@FrontCover-FrontColor", designDtl.FrontCover_FrontColor));
                OleDbParameters.Add(new OleDbParameter("@FrontCover-ObverseColor", designDtl.FrontCover_ObverseColor));
                OleDbParameters.Add(new OleDbParameter("@Content-FrontColor", designDtl.Content_FrontColor));
                OleDbParameters.Add(new OleDbParameter("@Content-ObverseColor", designDtl.Content_ObverseColor));
                OleDbParameters.Add(new OleDbParameter("@SpacialColor", designDtl.SpaciallColor));
                OleDbParameters.Add(new OleDbParameter("@SpacialColorCount", designDtl.SpacialColorCount));
                OleDbParameters.Add(new OleDbParameter("@AttachmentMemo", designDtl.AttachmentMemo));
                OleDbParameters.Add(new OleDbParameter("@AttachmentProvideDate", designDtl.AttachmentProvideDate));
                OleDbParameters.Add(new OleDbParameter("@CounterSing", designDtl.CounterSing));
                OleDbParameters.Add(new OleDbParameter("@CounterSingType", designDtl.CounterSingType));
                OleDbParameters.Add(new OleDbParameter("@CounterSingMemo", designDtl.CounterSingMemo));
                OleDbParameters.Add(new OleDbParameter("@Status", designDtl.Status));
                OleDbParameters.Add(new OleDbParameter("@InitalDate", designDtl.InitalDate));
                OleDbParameters.Add(new OleDbParameter("@InitalWorkTime", designDtl.InitalWorkTime));
                OleDbParameters.Add(new OleDbParameter("@InitalMemo", designDtl.InitalMemo));
                OleDbParameters.Add(new OleDbParameter("@FirstDate", designDtl.FinishDate));
                OleDbParameters.Add(new OleDbParameter("@FirstWorkTime", designDtl.FirstWorkTime));
                OleDbParameters.Add(new OleDbParameter("@FirstMemo", designDtl.FirstMemo));
                OleDbParameters.Add(new OleDbParameter("@SecondDate", designDtl.SecondDate));
                OleDbParameters.Add(new OleDbParameter("@SecondWorkTime", designDtl.SecondWorkTime));
                OleDbParameters.Add(new OleDbParameter("@SecondMemo", designDtl.SecondMemo));
                OleDbParameters.Add(new OleDbParameter("@ThirdDate", designDtl.ThirdWorkTime));
                OleDbParameters.Add(new OleDbParameter("@ThirdWorkTime", designDtl.ThirdWorkTime));
                OleDbParameters.Add(new OleDbParameter("@ThirdMemo", designDtl.ThirdMemo));
                OleDbParameters.Add(new OleDbParameter("@FinishDate", designDtl.FinishDate));
                OleDbParameters.Add(new OleDbParameter("@Acceptance", designDtl.Acceptance));
                OleDbParameters.Add(new OleDbParameter("@Designer", designDtl.Designer));

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

        public int DesignDtlUpdate(DesignDetailDataModel designDtl)
        {
            int result = 0;
            string sql = ""; 

            try 
            {
                sql = sql + " UPDATE [dbo].[DesignRequestDetail] Set";
                sql = sql + " [Type] = ?,";
                sql = sql + " [Count] = ?,";
                sql = sql + " [Layout-High] = ?,";
                sql = sql + " [Layout-Width] = ?,";
                sql = sql + " [Layout-Unit] = ?,";
                sql = sql + " [Layout-Type] = ?,";
                sql = sql + " [Layout-Page] = ?,";
                sql = sql + " [FrontCover] = ?,";
                sql = sql + " [FrontCoverType] = ?,";
                sql = sql + " [Content] = ?,";
                sql = sql + " [ContentType] = ?,";
                sql = sql + " [FrontCover-FrontColor] = ?,";
                sql = sql + " [FrontCover-ObverseColor] = ?,";
                sql = sql + " [Content-FrontColor] = ?,";
                sql = sql + " [Content-ObverseColor] = ?,";
                sql = sql + " [SpacialColor] = ?,";
                sql = sql + " [SpacialColorCount] = ?,";
                sql = sql + " [Request] = ?,";
                sql = sql + " [AttachmentMemo] = ?,";
                sql = sql + " [AttachmentProvideDate] = ?,";
                sql = sql + " [CounterSing] = ?,";
                sql = sql + " [CounterSingType] = ?,";
                sql = sql + " [CounterSingMemo] = ?,";
                sql = sql + " [Status] = ?,";
                sql = sql + " [InitalDate] = ?,";
                sql = sql + " [InitalWorkTime] = ?,";
                sql = sql + " [InitalMemo] = ?,";
                sql = sql + " [FirstDate] = ?,";
                sql = sql + " [FirstWorkTime] = ?,";
                sql = sql + " [FirstMemo] = ?,";
                sql = sql + " [SecondDate] = ?,";
                sql = sql + " [SecondWorkTime] = ?,";
                sql = sql + " [SecondMemo] = ?,";
                sql = sql + " [ThirdDate] = ?,";
                sql = sql + " [ThirdWorkTime] = ?,";
                sql = sql + " [ThirdMemo] = ?,";
                sql = sql + " [FinishDate] = ?,";
                sql = sql + " [Acceptance] = ?,";
                sql = sql + " [Designer]= ?";
                sql = sql + " WHERE  ID = ? AND MainID = ?";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@Type", designDtl.Type));
                OleDbParameters.Add(new OleDbParameter("@Count", designDtl.Count));
                OleDbParameters.Add(new OleDbParameter("@Layout-High", designDtl.Layout_High));
                OleDbParameters.Add(new OleDbParameter("@Layout-Width", designDtl.Layout_Width));
                OleDbParameters.Add(new OleDbParameter("@Layout-Unit", designDtl.Layout_Unit));
                OleDbParameters.Add(new OleDbParameter("@Layout-Type", designDtl.Layout_Type));
                OleDbParameters.Add(new OleDbParameter("@Layout-Page", designDtl.Layout_Page));
                OleDbParameters.Add(new OleDbParameter("@FrontCover", designDtl.FrontCover));
                OleDbParameters.Add(new OleDbParameter("@FrontCoverType", designDtl.FrontCoverType));
                OleDbParameters.Add(new OleDbParameter("@ContentType", designDtl.ContentType));
                OleDbParameters.Add(new OleDbParameter("@FrontCover-FrontColor", designDtl.FrontCover_FrontColor));
                OleDbParameters.Add(new OleDbParameter("@FrontCover-ObverseColor", designDtl.FrontCover_ObverseColor));
                OleDbParameters.Add(new OleDbParameter("@Content-FrontColor", designDtl.Content_FrontColor));
                OleDbParameters.Add(new OleDbParameter("@Content-ObverseColor", designDtl.Content_ObverseColor));
                OleDbParameters.Add(new OleDbParameter("@SpacialColor", designDtl.SpaciallColor));
                OleDbParameters.Add(new OleDbParameter("@SpacialColorCount", designDtl.SpacialColorCount));
                OleDbParameters.Add(new OleDbParameter("@AttachmentMemo", designDtl.AttachmentMemo));
                OleDbParameters.Add(new OleDbParameter("@AttachmentProvideDate", designDtl.AttachmentProvideDate));
                OleDbParameters.Add(new OleDbParameter("@CounterSing", designDtl.CounterSing));
                OleDbParameters.Add(new OleDbParameter("@CounterSingType", designDtl.CounterSingType));
                OleDbParameters.Add(new OleDbParameter("@CounterSingMemo", designDtl.CounterSingMemo));
                OleDbParameters.Add(new OleDbParameter("@Status", designDtl.Status));
                OleDbParameters.Add(new OleDbParameter("@InitalDate", designDtl.InitalDate));
                OleDbParameters.Add(new OleDbParameter("@InitalWorkTime", designDtl.InitalWorkTime));
                OleDbParameters.Add(new OleDbParameter("@InitalMemo", designDtl.InitalMemo));
                OleDbParameters.Add(new OleDbParameter("@FirstDate", designDtl.FinishDate));
                OleDbParameters.Add(new OleDbParameter("@FirstWorkTime", designDtl.FirstWorkTime));
                OleDbParameters.Add(new OleDbParameter("@FirstMemo", designDtl.FirstMemo));
                OleDbParameters.Add(new OleDbParameter("@SecondDate", designDtl.SecondDate));
                OleDbParameters.Add(new OleDbParameter("@SecondWorkTime", designDtl.SecondWorkTime));
                OleDbParameters.Add(new OleDbParameter("@SecondMemo", designDtl.SecondMemo));
                OleDbParameters.Add(new OleDbParameter("@ThirdDate", designDtl.ThirdWorkTime));
                OleDbParameters.Add(new OleDbParameter("@ThirdWorkTime", designDtl.ThirdWorkTime));
                OleDbParameters.Add(new OleDbParameter("@ThirdMemo", designDtl.ThirdMemo));
                OleDbParameters.Add(new OleDbParameter("@FinishDate", designDtl.FinishDate));
                OleDbParameters.Add(new OleDbParameter("@Acceptance", designDtl.Acceptance));
                OleDbParameters.Add(new OleDbParameter("@Designer", designDtl.Designer));
                OleDbParameters.Add(new OleDbParameter("@ID", designDtl.ID));
                OleDbParameters.Add(new OleDbParameter("@MainID", designDtl.MainID));

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

        public int DesignDtlDelete(int id ,int mainid)
        {
            int result = 0;
            string sql = "";

            try
            {
                sql = "";
                sql = sql + " DELETE FROM DesignRequestDetail WHERE ID = ? AND MainID = ?";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@ID", id));
                OleDbParameters.Add(new OleDbParameter("@MainID", mainid));

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

        public int getDesignDtlMaxID(string mainid) 
        {
            int result = 0;
            string sql = "";
            DataTable dt = new DataTable();

            try 
            {
                sql = sql + " SELECT COUNT(ID) AS CNT FROM DesignRequestDetail WHERE MainID =? ";

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@MainID", mainid));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("DesignRequestDetail", sql, OleDbParameters);

                if (dt.Rows.Count != 0)
                {
                    result = int.Parse(dt.Rows[0]["CNT"].ToString().Trim()) + 1;
                }
                else 
                {
                    result = result + 1;
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

        public string getDatelimite(string type,string id) 
        {
            string result = "";
            string sql = "";
            DataTable dt = new DataTable();

            try 
            {
                if (type == "Main")
                {
                    sql = sql + " SELECT Datelimite FROM [dbo].[DesignRequest] WHERE ID =?";
                }
                else 
                {
                    sql = sql + " SELECT Datelimite FROM [dbo].[DesignRequest] WHERE ID =";
                    sql = sql + "   (SELECT TOP 1 MainID FROM [dbo].[DesignRequestDetail] WHERE ID = ?)";                
                }

                List<OleDbParameter> OleDbParameters = new List<OleDbParameter>();
                OleDbParameters.Add(new OleDbParameter("@ID", id));

                Info info = new Info();
                Command cmd = new Command(info);
                dt = cmd.GetDataTable("DesignRequest", sql,OleDbParameters);

                if (dt.Rows.Count > 0) 
                {
                    result = dt.Rows[0]["Datelimite"].ToString().Trim();
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
    }
}