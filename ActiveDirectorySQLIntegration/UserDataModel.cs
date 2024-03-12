namespace ActiveDirectorySQLIntegration
{
    public class UserDataModel
    {
        public required string objectSID { get; set; }
        public string? sAMAccountName { get; set; }
        public string? userPrincipalName {  get; set; }
        public string? displayName { get; set; }
        public string? department { get; set; }
        public string? company { get; set; }
        public string? description { get; set; }
        public string? title { get; set; }
    }
}
