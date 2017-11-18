using System.Data;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Collections.Generic;
using com.data.structure;

namespace service.designservice
{
    // 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码和配置文件中的接口名“IDesignService”。
    [ServiceContract]
    [ServiceKnownType(typeof(DesignDataModel))]
    [ServiceKnownType(typeof(List<DesignDataModel>))]
    [ServiceKnownType(typeof(DesignDetailDataModel))]
    [ServiceKnownType(typeof(List<DesignDetailDataModel>))]

    public interface IDesignService
    {
        [OperationContract]
        DataTable GetDesignList();

        [OperationContract]
        DataTable GetDesignDtlList(int mainid);

        [OperationContract]
        int DesignInsert(DesignDataModel design);

        [OperationContract]
        int DesignUpdate(DesignDataModel design);

        [OperationContract]
        int DesignDelete(int id);

        [OperationContract]
        int DesignDtlInsert(DesignDetailDataModel designDtl);

        [OperationContract]
        int DesignDtlUpdate(DesignDetailDataModel designDtl);

        [OperationContract]
        int DesignDtlDelete(int id ,int mainid);

        [OperationContract]
        int getDesignDtlMaxID(string mainid);

        [OperationContract]
        string getDatelimite(string type,string id);
    }
}