using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

// 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码和配置文件中的接口名“IUserSvc”。

    [ServiceContract]
    public interface IUserSvc
    {
        [OperationContract]
        [WebInvoke(Method = "POST", BodyStyle = WebMessageBodyStyle.WrappedRequest, ResponseFormat = WebMessageFormat.Json)]

        //ResultMessage SaveUserData(NameValuePairs nvps);

        ResultMessage checkAD_Account(NameValuePair nvps);
    }

    [DataContract]
    public class NameValuePair
    {
        [DataMember]
        public string name { get; set; }
        [DataMember]
        public string value { get; set; }
    }

    [DataContract]
    public class ResultMessage
    {
        [DataMember]
        public string Message { get; set; }
    }

    [CollectionDataContract(Namespace = "")]
    public class NameValuePairs : List<NameValuePair>
    {
    }
