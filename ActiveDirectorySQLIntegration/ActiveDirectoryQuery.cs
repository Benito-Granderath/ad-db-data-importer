using ActiveDirectorySQLIntegration;
using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Globalization;
using System.Security.Principal;

#pragma warning disable CA1416
public static class ActiveDirectoryService
{
    public static List<UserDataModel> users = new();
    public static List<GroupDataModel> groups = new();

    public static void ActiveDirectoryQuery()
    {
        using (var context = new PrincipalContext(ContextType.Domain))
        {
            using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
            {

                foreach (var result in searcher.FindAll())
                {
                    DirectoryEntry? de = result.GetUnderlyingObject() as DirectoryEntry;

                    var objectSidProperty = de.Properties["objectSid"].Value as byte[];
                    if (objectSidProperty != null)
                    {

                        var user = new UserDataModel
                        {
                            objectSID = new SecurityIdentifier(objectSidProperty, 0).ToString(),
                            sAMAccountName = de.Properties["sAMAccountName"].Value?.ToString() ?? string.Empty,
                            userPrincipalName = de.Properties["userPrincipalName"].Value?.ToString() ?? string.Empty,
                            displayName = de.Properties["displayName"].Value?.ToString() ?? string.Empty,
                            department = de.Properties["department"].Value?.ToString() ?? string.Empty,
                            company = de.Properties["company"].Value?.ToString() ?? string.Empty,
                            description = de.Properties["description"].Value?.ToString() ?? string.Empty,
                            title = de.Properties["title"].Value?.ToString() ?? string.Empty,
                        };
                        users.Add(user);
                        Console.WriteLine(de.Properties["sAMAccountName"].Value?.ToString());
                        Console.WriteLine(de.Properties["department"].Value?.ToString());
                    }
                }
            }

            using (var searcher = new PrincipalSearcher(new GroupPrincipal(context)))
            {
                foreach (var result in searcher.FindAll())
                {
                    DirectoryEntry? de = result.GetUnderlyingObject() as DirectoryEntry;

                    var objectSidProperty = de.Properties["objectSid"].Value as byte[];
                    if (objectSidProperty != null)
                    {
                        var group = new GroupDataModel
                        {
                            objectSID = new SecurityIdentifier(objectSidProperty, 0).ToString(),
                            description = de.Properties["description"].Value?.ToString() ?? string.Empty,
                            groupType = de.Properties["groupType"].Value?.ToString() ?? string.Empty,
                            ManagedBy = de.Properties["ManagedBy"].Value?.ToString() ?? string.Empty,
                            WhenCreated = DateTime.ParseExact(de.Properties["whenCreated"].Value?.ToString() ?? string.Empty, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal),
                            WhenChanged = DateTime.ParseExact(de.Properties["whenChanged"].Value?.ToString() ?? string.Empty, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal)
                            
                        };
                        groups.Add(group);
                        Console.WriteLine(de.Properties["description"].Value?.ToString());
                        Console.WriteLine(de.Properties["groupType"]);
                    }
                }
            }
            Console.WriteLine(groups);

            Console.WriteLine(users);
        }
    }
}
