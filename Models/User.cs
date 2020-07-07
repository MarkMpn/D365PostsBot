using System;
using Microsoft.Azure.Cosmos.Table;

namespace MarkMpn.D365PostsBot.Models
{
    public class User : TableEntity
    {
        public User(string username) : base(username, "")
        {
        }

        public User()
        {
        }

        public string UserId { get; set; }

        public string TenantId { get; set; }

        public string ServiceUrl { get; set; }

        public string LastDomainName { get; set; }

        public Guid? LastPostId { get; set; }
    }
}
