using ActiveDirectorySQLIntegration;
using System.Configuration;
using System.Data.SqlClient;
namespace ActiveDirectorySQLIntegration
{
    public class ServerData { 
        public static void Main()
        {
            ActiveDirectoryService.ActiveDirectoryQuery();
            string connectionString = "Server=dc1-db01;Database=wsmb;Trusted_Connection=True;";

            foreach (var user in ActiveDirectoryService.users)
            {
                using (var connection = new SqlConnection(connectionString))
                { 
                    string userQuery = "INSERT INTO ActiveDirectoryUserData (objectSID, sAMAccountName, userPrincipalName, displayName, department, company, description, title) VALUES (@objectSid, @sAMAccountName, @userPrincipalName, @displayName, @department, @company, @description, @title)";

                    using (var command = new SqlCommand(userQuery, connection))
                    {
                        command.Parameters.AddWithValue("@objectSid", user.objectSID);
                        command.Parameters.AddWithValue("@sAMAccountName", user.sAMAccountName);
                        command.Parameters.AddWithValue("@userPrincipalName", user.userPrincipalName);
                        command.Parameters.AddWithValue("@displayName", user.displayName);
                        command.Parameters.AddWithValue("@department", user.department);
                        command.Parameters.AddWithValue("@company", user.company);
                        command.Parameters.AddWithValue("@description", user.description);
                        command.Parameters.AddWithValue("@title", user.title);

                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
            foreach (var group in ActiveDirectoryService.groups)
            {
                using (var connection = new SqlConnection(connectionString))
                {
                    string groupQuery = "INSERT INTO ActiveDirectoryGroupData (objectSID, description, groupType, ManagedBy, WhenCreated, WhenChanged) VALUES (@objectSID, @description, @groupType, @ManagedBy, @WhenCreated, @WhenChanged)";
                
                    using (var command = new SqlCommand (groupQuery, connection)) 
                    {
                        command.Parameters.AddWithValue("@objectSID", group.objectSID);
                        command.Parameters.AddWithValue("@description", group.description);
                        command.Parameters.AddWithValue("@groupType", group.groupType);
                        command.Parameters.AddWithValue("@ManagedBy", group.ManagedBy);
                        command.Parameters.AddWithValue("@WhenCreated", group.WhenCreated);
                        command.Parameters.AddWithValue("@WhenChanged", group.WhenChanged);

                        connection.Open();
                        command.ExecuteNonQuery();
                    }
                }

            }
        }
    }
}
