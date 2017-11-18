using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using System.Text;


    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    public class UserSvc : IUserSvc
    {

        public ResultMessage checkAD_Account(NameValuePair nvps)
        {
            var result = new ResultMessage() { Message = "User saved successfully." };


            return result;
        }


        public ResultMessage SaveUserData(NameValuePairs nvps)
        {
            var columns = new StringBuilder();
            var parameters = new StringBuilder();
            var rm = new ResultMessage() { Message = "User saved successfully." };

            //using (var conn = new SqlConnection(Properties.Settings.Default.AppDb))
            //using (var cmd = new SqlCommand() { CommandType = CommandType.Text, Connection = conn })
            //{
            //    var paramCollection = cmd.Parameters;
            //    var rows = 0;

            //    try
            //    {
            //        if (nvps == null || nvps.Count == 0)
            //        {
            //            throw new ArgumentNullException("nvps");
            //        }

            //        foreach (NameValuePair nvp in nvps)
            //        {
            //            columns.Append(nvp.name).Append(",");
            //            parameters.Append("@").Append(nvp.name).Append(",");
            //            paramCollection.Add(new SqlParameter("@" + nvp.name, nvp.value));
            //        }

            //        // KLUDGE: Remove trailing commas.
            //        columns.Remove(columns.Length - 1, 1);
            //        parameters.Remove(parameters.Length - 1, 1);
            //        cmd.CommandText = String.Format(Properties.Settings.Default.InsertSql, columns.ToString(), parameters.ToString());
            //        conn.Open();
            //        rows = cmd.ExecuteNonQuery();
            //        rm.Message = String.Format(Properties.Settings.Default.SuccessMessage, rows, cmd.CommandText);
            //    }
            //    catch (Exception ex)
            //    {
            //        rm.Message = ex.Message;
            //    }
            //}

            return rm;
        }
    }

