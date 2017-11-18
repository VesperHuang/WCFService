using System.Data;
using System.Runtime.Serialization;
using System.ServiceModel;

using com.data.structure;

// 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码和配置文件中的接口名“IService”。
[ServiceContract]
[ServiceKnownType(typeof(AttachmentDataModel))]

public interface IService
{
    [OperationContract]
    DataTable GetCheckUserAccount(string systemcode,string account, string password);

    [OperationContract]
    DataTable getUserArea(string userNo);

    [OperationContract]
    DataTable getUserMenuRule(string groupID, string userNo);

    [OperationContract]
    DataTable getMenu(string School_TypeID);

    [OperationContract]
    DataTable getMenuByGroup(string School_TypeID, string User_GroupID);

    [OperationContract]
    DataTable getArea_Branch();

    [OperationContract]
    DataTable getSystem_Code(string Sys_Code = "");

    [OperationContract]
    DataTable getGroupRule(string SchoolType, string GroupID);

    [OperationContract]
    int setGroupRule(string SchoolType, string GroupID, string[] Menu_Code);

    [OperationContract]
    string getUserInfo(string User_NO = "", string User_EnName = "");

    [OperationContract]
    DataTable getBranchRule(string User_No);

    [OperationContract]
    int setBranchRule(string userno, string branchs, string type);

    [OperationContract]
    int InsertSysCode(string input);

    [OperationContract]
    int UpdateSysCode(string input);

    [OperationContract]
    int DeleteSysCode(string input);

    [OperationContract]
    DataTable getUserTable(string input = "");

    [OperationContract]
    int InsertUserTable(string input);

    [OperationContract]
    int UpdateUserTable(string input);

    [OperationContract]
    int DeleteUserTable(string userno);

    [OperationContract]
    DataTable getUserBranchs(string input);

    [OperationContract]
    string getUserBranchsByString(string schooltypeid, string userid);

    [OperationContract]
    DataTable getMenu_Rule(string input);

    [OperationContract]
    int setMenu_Rule(string input);

    [OperationContract]
    DataTable getUserBranchsBySchoolType(string input);


    [OperationContract]
    DataTable GetUserData(string input = "");


    [OperationContract]
    string getUrlFromJdboarweb(string input);

    [OperationContract]
    int InsertAttachment(AttachmentDataModel attachment);


    [OperationContract]
    string getSystemCode_ItemName(string input);


    [OperationContract]
    int DeleteAttachment(string Category, string fromid);


    [OperationContract]
    bool IsAttachment(string Category, string fromid);
}


// 使用下面示例中说明的数据约定将复合类型添加到服务操作。
[DataContract]
public class CompositeType
{
    bool boolValue = true;
    string stringValue = "Hello ";

    [DataMember]
    public bool BoolValue
    {
        get { return boolValue; }
        set { boolValue = value; }
    }

    [DataMember]
    public string StringValue
    {
        get { return stringValue; }
        set { stringValue = value; }
    }
}
