using ActiveDirectorySQLIntegration;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics.Tracing;
using System.Text.RegularExpressions;
namespace ActiveDirectorySQLIntegration
{
    public class ServerData
    {
        string connectionString = "Server=dc1-db01;Database=wsmb;Trusted_Connection=True;";
        public static void Main()
        {
            ActiveDirectoryService.ActiveDirectoryQuery();
            var SD = new ServerData();
            int userLength = ActiveDirectoryService.users.Count;
            int currentUserLength = 0;
            foreach (var user in ActiveDirectoryService.users)
            {
                using (SqlConnection conn = new SqlConnection(SD.connectionString))
                {
                    conn.Open();

                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.Parameters.AddWithValue("@objectSid", user.objectSID);
                    cmd.Parameters.AddWithValue("@sAMAccountName", user.sAMAccountName);
                    cmd.Parameters.AddWithValue("@userPrincipalName", user.userPrincipalName);
                    cmd.Parameters.AddWithValue("@displayName", user.displayName);
                    cmd.Parameters.AddWithValue("@TelephoneNumber", user.TelephoneNumber);
                    cmd.Parameters.AddWithValue("@department", user.department);
                    cmd.Parameters.AddWithValue("@company", user.company);
                    cmd.Parameters.AddWithValue("@description", user.description);
                    cmd.Parameters.AddWithValue("@title", user.title);

                    cmd.CommandText = "SELECT COUNT(1) FROM AD_USER WHERE objectSID = @objectSid";
                    int count = Convert.ToInt32(cmd.ExecuteScalar());

                    if (count == 0)
                    {
                        cmd.CommandText = "INSERT INTO AD_USER (objectSID, sAMAccountName, userPrincipalName, displayName, TelephoneNumber, department, company, description, title) VALUES (@objectSid, @sAMAccountName, @userPrincipalName, @displayName, @TelephoneNumber, @department, @company, @description, @title)";
                        Console.WriteLine($"User Eintrag {currentUserLength} von {userLength} geschrieben");
                    }
                    else
                    {
                        cmd.CommandText = "UPDATE AD_USER SET sAMAccountName = @sAMAccountName, userPrincipalName = @userPrincipalName, displayName = @displayName, TelephoneNumber = @TelephoneNumber, department = @department, company = @company, description = @description, title = @title WHERE objectSID = @objectSid";
                        Console.WriteLine($"User Eintrag {currentUserLength} von {userLength} aktualisiert");
                    }

                    cmd.ExecuteNonQuery();
                    currentUserLength++;
                }
            }
            using (var connection = new SqlConnection(SD.connectionString))
            {
                string departmentQuery = "MERGE INTO AD_DEPARTMENT AS target\r\nUSING (SELECT DISTINCT department FROM AD_USER) AS source\r\nON target.departmentName = source.department\r\nWHEN MATCHED THEN \r\n    UPDATE SET target.departmentName = source.department\r\nWHEN NOT MATCHED BY TARGET THEN \r\n    INSERT (departmentName) VALUES (source.department);";
                using (var command = new SqlCommand(departmentQuery, connection))
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    Console.WriteLine("Abteilungen gelesen und hinzuefügt!");
                }
            }

            int groupLength = ActiveDirectoryService.groups.Count;
            int currentGroupLength = 0;
            foreach (var group in ActiveDirectoryService.groups)
            {
                using (SqlConnection conn = new SqlConnection(SD.connectionString))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn;
                    cmd.Parameters.AddWithValue("@objectSID", group.objectSID);
                    cmd.Parameters.AddWithValue("@CommonName", group.CommonName);
                    cmd.Parameters.AddWithValue("@description", group.description);
                    cmd.Parameters.AddWithValue("@groupType", group.groupType);
                    cmd.Parameters.AddWithValue("@ManagedBy", group.ManagedBy);
                    cmd.Parameters.AddWithValue("@WhenCreated", group.WhenCreated);
                    cmd.Parameters.AddWithValue("@WhenChanged", group.WhenChanged);
                    cmd.CommandText = "SELECT COUNT(1) FROM AD_GROUP WHERE objectSID = @objectSID";
                    int count = Convert.ToInt32(cmd.ExecuteScalar());
                    if (count == 0) 
                    {
                        cmd.CommandText = "INSERT INTO AD_GROUP (objectSID, CommonName, description, groupType, ManagedBy, WhenCreated, WhenChanged) VALUES (@objectSID, @CommonName, @description, @groupType, @ManagedBy, @WhenCreated, @WhenChanged)";
                        Console.WriteLine($"Gruppen Eintrag {currentGroupLength} von {groupLength} geschrieben");
                    }
                    else
                    {
                        cmd.CommandText = "UPDATE AD_GROUP SET CommonName = @CommonName, description = @description, groupType = @groupType, ManagedBy = @ManagedBy, WhenCreated = @WhenCreated, WhenChanged = @WhenChanged WHERE objectSID = @objectSID";
                        Console.WriteLine($"Gruppen Eintrag {currentGroupLength} von {groupLength} aktualisiert");
                    }
                    cmd.ExecuteNonQuery();
                    currentGroupLength++;
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

                            string checkQuery = @"SELECT COUNT(1) FROM AD_USER_GROUP_JUNCTION WHERE UserSID = @UserSID AND GroupSID = @GroupSID";
                            SqlCommand checkCommand = new SqlCommand(checkQuery, connection);
                            checkCommand.Parameters.AddWithValue("@UserSID", user.objectSID);
                            checkCommand.Parameters.AddWithValue("@GroupSID", groupSid);

                            int count = Convert.ToInt32(checkCommand.ExecuteScalar());

                            if (count == 0)
                            {
                                var insertCommand = new SqlCommand(@"INSERT INTO AD_USER_GROUP_JUNCTION (UserSID, GroupSID, sAMAccountName, groupCommonName) VALUES (@UserSID, @GroupSID, @sAMAccountName, @groupCommonName)", connection);
                                insertCommand.Parameters.AddWithValue("@UserSID", user.objectSID);
                                insertCommand.Parameters.AddWithValue("@GroupSID", groupSid);
                                insertCommand.Parameters.AddWithValue("@sAMAccountName", user.sAMAccountName);
                                insertCommand.Parameters.AddWithValue("@groupCommonName", groupCommonName);

                                try
                                {
                                    insertCommand.ExecuteNonQuery();
                                    Console.WriteLine($"Junction Eintrag für Nutzer {user.sAMAccountName} und Gruppe {groupCommonName} geschrieben.");
                                }
                                catch (SqlException ex)
                                {
                                    Console.WriteLine($"Fehler: {ex.Message}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Junction Eintrag für Nutzer {user.sAMAccountName} und Gruppe {groupCommonName} existiert bereits.");
                            }
                        }
                    }
                }
            }
        }

        public static string GetGroupCommonNameBySid(string groupSid, SqlConnection connection)
        {
            string query = "SELECT CommonName FROM AD_GROUP WHERE objectSID = @GroupSID";
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@GroupSID", groupSid);
                var result = command.ExecuteScalar();
                return result != null ? result.ToString() : "Unbekannte Gruppe";
            }
        }


    }
}
