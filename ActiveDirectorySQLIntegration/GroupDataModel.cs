namespace ActiveDirectorySQLIntegration
{
    public class GroupDataModel
    {
        public required string objectSID { get; set; }
        public string? description { get; set; }
        public string? groupType { get; set; }
        public string? ManagedBy { get; set; }
        public DateTime WhenCreated { get; set; }
        public DateTime WhenChanged { get; set; }

    }
}
