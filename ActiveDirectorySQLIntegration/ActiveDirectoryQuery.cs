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
    public static string? GetGroupObjectSIDByDN(string distinguishedName)
    {
        try
        {
            string ldapPath = "LDAP://wuensche-group.local";
            using (DirectoryEntry de = new DirectoryEntry(ldapPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(de))
                {
                    searcher.Filter = $"(distinguishedName={distinguishedName})";
                    searcher.PropertiesToLoad.Add("objectSid");

                    SearchResult result = searcher.FindOne();
                    if (result != null)
                    {
                        byte[]? sidBytes = result.Properties["objectSid"][0] as byte[];
                        if (sidBytes != null)
                        {
                            return new SecurityIdentifier(sidBytes, 0).ToString();
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fehler {distinguishedName}: {ex.Message}");
        }
        return null;
    }
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
                            TelephoneNumber = de.Properties["TelephoneNumber"].Value?.ToString() ?? string.Empty,
                            department = de.Properties["department"].Value?.ToString() ?? string.Empty,
                            company = de.Properties["company"].Value?.ToString() ?? string.Empty,
                            description = de.Properties["description"].Value?.ToString() ?? string.Empty,
                            title = de.Properties["title"].Value?.ToString() ?? string.Empty,
                        };

                        users.Add(user);
                        Console.WriteLine($"Nutzer {user.sAMAccountName} zu Liste hinzugefügt");
                        var memberOfGroups = de.Properties["memberOf"];
                        if (memberOfGroups != null)
                        {
                            foreach (var dn in memberOfGroups)
                            {
                                string groupDn = dn.ToString();
                                string groupSid = GetGroupObjectSIDByDN(groupDn);
                                user.GroupSIDs.Add(groupSid);
                                
                            }
                        }
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
                            CommonName = de.Properties["cn"].Value?.ToString() ?? string.Empty,
                            description = de.Properties["description"].Value?.ToString() ?? string.Empty,
                            groupType = de.Properties["groupType"].Value?.ToString() ?? string.Empty,
                            ManagedBy = de.Properties["ManagedBy"].Value?.ToString() ?? string.Empty,
                            WhenCreated = DateTime.ParseExact(de.Properties["whenCreated"].Value?.ToString() ?? string.Empty, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal),
                            WhenChanged = DateTime.ParseExact(de.Properties["whenChanged"].Value?.ToString() ?? string.Empty, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal)
                            
                        };
                        groups.Add(group);
                        Console.WriteLine($"Gruppe {group.CommonName} zu Liste hinzugefügt");
                    }
                }
            }
        }
    }
   
}
