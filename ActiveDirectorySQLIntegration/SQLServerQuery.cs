using ActiveDirectorySQLIntegration;
using System.Configuration;
using System.Data.SqlClient;
namespace ActiveDirectorySQLIntegration
{
    public class ServerData
    {
        string connectionString = "Server=dc1-db01;Database=wsmb;Trusted_Connection=True;";
        public static void Main()
        {
            ActiveDirectoryService.ActiveDirectoryQuery();
            var SD = new ServerData();
            foreach (var user in ActiveDirectoryService.users)
            {
                using (var connection = new SqlConnection(SD.connectionString))
                {
                    string userQuery = "INSERT INTO ADUser (objectSID, sAMAccountName, userPrincipalName, displayName, TelephoneNumber, department, company, description, title) VALUES (@objectSid, @sAMAccountName, @userPrincipalName, @displayName, @TelephoneNumber, @department, @company, @description, @title)";

                    using (var command = new SqlCommand(userQuery, connection))
                    {
                        command.Parameters.AddWithValue("@objectSid", user.objectSID);
                        command.Parameters.AddWithValue("@sAMAccountName", user.sAMAccountName);
                        command.Parameters.AddWithValue("@userPrincipalName", user.userPrincipalName);
                        command.Parameters.AddWithValue("@displayName", user.displayName);
                        command.Parameters.AddWithValue("@TelephoneNumber", user.TelephoneNumber);
                        command.Parameters.AddWithValue("@department", user.department);
                        command.Parameters.AddWithValue("@company", user.company);
                        command.Parameters.AddWithValue("@description", user.description);
                        command.Parameters.AddWithValue("@title", user.title);

                        connection.Open();
                        command.ExecuteNonQuery();
                        Console.WriteLine($"User Eintrag geschrieben");
                    }
                }
            }
            foreach (var group in ActiveDirectoryService.groups)
            {
                using (var connection = new SqlConnection(SD.connectionString))
                {
                    string groupQuery = "INSERT INTO ADGroup (objectSID, CommonName, description, groupType, ManagedBy, WhenCreated, WhenChanged) VALUES (@objectSID, @CommonName, @description, @groupType, @ManagedBy, @WhenCreated, @WhenChanged)";
                    using (var command = new SqlCommand(groupQuery, connection))
                    {
                        command.Parameters.AddWithValue("@objectSID", group.objectSID);
                        command.Parameters.AddWithValue("@CommonName", group.CommonName);
                        command.Parameters.AddWithValue("@description", group.description);
                        command.Parameters.AddWithValue("@groupType", group.groupType);
                        command.Parameters.AddWithValue("@ManagedBy", group.ManagedBy);
                        command.Parameters.AddWithValue("@WhenCreated", group.WhenCreated);
                        command.Parameters.AddWithValue("@WhenChanged", group.WhenChanged);

                        connection.Open();
                        command.ExecuteNonQuery();
                        Console.WriteLine($"Gruppen Eintrag geschrieben");
                    }
                }
            }
            InsertUserGroupMemberships(SD.connectionString);
        }
        public static void InsertUserGroupMemberships(string connectionString)
        {
            foreach (var user in ActiveDirectoryService.users)
            {
                foreach (var groupSid in user.GroupSIDs)
                {
                    if (!string.IsNullOrEmpty(groupSid))
                    {
                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            string groupCommonName = GetGroupCommonNameBySid(groupSid, connection);
                            var command = new SqlCommand(@"INSERT INTO ADUserGroupJunction (UserSID, GroupSID, sAMAccountName, groupCommonName) VALUES (@UserSID, @GroupSID, @sAMAccountName, @groupCommonName)", connection);
                            command.Parameters.AddWithValue("@UserSID", user.objectSID);
                            command.Parameters.AddWithValue("@GroupSID", groupSid);
                            command.Parameters.AddWithValue("@sAMAccountName", user.sAMAccountName);
                            command.Parameters.AddWithValue("@groupCommonName", groupCommonName);
                            try
                            {
                                command.ExecuteNonQuery();
                                Console.WriteLine($"Junction Eintrag geschrieben");
                            }
                            catch (SqlException ex)
                            {
                                Console.WriteLine($"Fehler {ex.Message}");
                            }
                        }
                    }
                }
            }
        }
        public static string GetGroupCommonNameBySid(string groupSid, SqlConnection connection)
        {
            string query = "SELECT CommonName FROM ADGroup WHERE objectSID = @GroupSID";
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@GroupSID", groupSid);
                var result = command.ExecuteScalar();
                return result != null ? result.ToString() : "Unbekannte Gruppe";
            }
        }


    }
}
