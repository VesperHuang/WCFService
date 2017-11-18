using System.Data;
using System.ServiceModel;


// 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码和配置文件中的接口名“IReportService”。
[ServiceContract]
public interface IReportService
{
    [OperationContract]
    int InsertSchool_Students_Status_Count(string input);

    [OperationContract]
    DataTable getSchool_Students_Status_Count(string input);

    [OperationContract]
    DataTable getRpt0001(string input = "");

    [OperationContract]
    int setReportStatus(string input);

    [OperationContract]
    bool chkRpt00001(string input);

    [OperationContract]
    int InsertRpt00001(string input);

    [OperationContract]
    int UpdateRpt00001(string input);

    [OperationContract]
    int DeleteRpt00001(string input);

    [OperationContract]
    int InsertReportUpload(string input);

    [OperationContract]
    bool chkReportUpload(string input);

    [OperationContract]
    DataTable getReportUpload(string input);

    [OperationContract]
    int DeleteReportUpload(string input);
}
