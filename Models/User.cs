using System;
using Azure;
using Azure.Data.Tables;

namespace MarkMpn.D365PostsBot.Models
{
    public class User : ITableEntity
    {
        public User(string username)
        {
            PartitionKey = username;
            RowKey = "";
        }

        public User()
        {
        }

        public string UserId { get; set; }

        public string TenantId { get; set; }

        public string ServiceUrl { get; set; }

        public string LastDomainName { get; set; }

        public Guid? LastPostId { get; set; }

        public string PartitionKey { get; set; }

        public string RowKey { get; set; }

        public DateTimeOffset? Timestamp { get; set; }

        public ETag ETag { get; set; }
    }
}
