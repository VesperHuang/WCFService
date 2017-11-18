using com.data.structure;
using System.Collections.Generic;
using System.ServiceModel;

// 注意: 使用“重构”菜单上的“重命名”命令，可以同时更改代码和配置文件中的接口名“IMeetingRoomService”。
[ServiceContract]
[ServiceKnownType(typeof(AppointmentInfo))]
[ServiceKnownType(typeof(List<AppointmentInfo>))]
[ServiceKnownType(typeof(RoomInfo))]
[ServiceKnownType(typeof(List<RoomInfo>))]
[ServiceKnownType(typeof(UserInfo))]
[ServiceKnownType(typeof(List<UserInfo>))]

public interface IMeetingRoomService
{

    [OperationContract]
    [ServiceKnownType(typeof(List<AppointmentInfo>))]
    List<AppointmentInfo> GetAppointments();

    [OperationContract]
    [ServiceKnownType(typeof(List<RoomInfo>))]
    List<RoomInfo> GetMeetingRooms();

    [OperationContract]
    [ServiceKnownType(typeof(List<UserInfo>))]
    List<UserInfo> GetUser();

    [OperationContract]
    [ServiceKnownType(typeof(AppointmentInfo))]
    int updateAppointment(AppointmentInfo app);


    [OperationContract]
    [ServiceKnownType(typeof(AppointmentInfo))]
    int insertAppointment(AppointmentInfo app);


    [OperationContract]
    int deleteAppointment(string id);
}
