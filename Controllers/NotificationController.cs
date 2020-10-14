using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using AdaptiveCards;
using MarkMpn.D365PostsBot.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Cosmos.Table;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using Microsoft.PowerPlatform.Cds.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json.Linq;
using Entity = Microsoft.Xrm.Sdk.Entity;

namespace MarkMpn.D365PostsBot.Controllers
{
    [Route("api/notification")]
    [ApiController]
    public class NotificationController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ConcurrentDictionary<string, ConcurrentDictionary<string, EntityMetadata>> _metadata;

        static NotificationController()
        {

        }

        public NotificationController(IConfiguration config, ConcurrentDictionary<string, ConcurrentDictionary<string, EntityMetadata>> metadata)
        {
            _config = config;
            _metadata = metadata;
        }

        private string DomainName => Request.Headers["x-ms-dynamics-organization"].Single();

        [HttpPost]
        public async Task<IActionResult> PostAsync([FromQuery] string code, [FromBody] JObject requestContext)
        {
            if (code != _config.GetValue<string>("WebhookKey"))
                return Unauthorized();

            try
            {
                // Get a reference to the post/postcomment that has just been created. The requestContext holds a RemoteExecutionContext
                // object, so extract out the PrimaryEntityName and PrimaryEntityId properties
                var postReference = new EntityReference(requestContext.Value<string>("PrimaryEntityName"), new Guid(requestContext.Value<string>("PrimaryEntityId")));

                using (var org = new CdsServiceClient(new Uri("https://" + DomainName), _config.GetValue<string>("MicrosoftAppId"), _config.GetValue<string>("MicrosoftAppPassword"), true, null))
                {
                    var post = org.Retrieve(postReference.LogicalName, postReference.Id, new ColumnSet(true));
                    var postComment = post;

                    if (postReference.LogicalName == "postcomment")
                    {
                        // This is a comment to an existing post, go and retrieve the original post
                        post = org.Retrieve("post", postComment.GetAttributeValue<EntityReference>("postid").Id, new ColumnSet(true));
                    }

                    // Get the entity the post is on
                    var entityRef = post.GetAttributeValue<EntityReference>("regardingobjectid");
                    var entity = org.Retrieve(entityRef.LogicalName, entityRef.Id, new ColumnSet(true));

                    post = GetFullPostText(org, entityRef, post, ref postComment);

                    // Get the users to notify
                    var entityRelationships = new Dictionary<Guid, Link>();
                    var usersToNotify = new HashSet<EntityReference>();

                    if (postComment != post)
                    {
                        entityRelationships.Add(postComment.Id, new Link { From = post.ToEntityReference(), Description = "Contains Comment" });
                        entityRelationships.Add(post.Id, new Link { From = postComment.ToEntityReference(), Description = "Is Comment On" });
                    }

                    entityRelationships.Add(entity.Id, new Link { From = post.ToEntityReference(), Description = "Is Posted On" });

                    if (entity.LogicalName == "systemuser")
                        usersToNotify.Add(entity.ToEntityReference());

                    GetInterestedUsers(post, postComment, entity, org, usersToNotify, entityRelationships);
                    ExpandTeamsToUsers(usersToNotify, org, entityRelationships);
                    AddUserFollows(usersToNotify, org, entityRelationships);

                    // Remove the user who added the post
                    usersToNotify.Remove(postComment.GetAttributeValue<EntityReference>("createdby"));

                    // Remove any possible null value
                    usersToNotify.Remove(null);

                    if (usersToNotify.Any())
                    {
                        string avatarUrl = null;

                        // Connect to storage account so we can look up the details we need to send the message to each user
                        var connectionString = _config.GetConnectionString("Storage");
                        var storageAccount = CloudStorageAccount.Parse(connectionString);
                        var tableClient = storageAccount.CreateCloudTableClient();
                        var table = tableClient.GetTableReference("users");

                        foreach (var userRef in usersToNotify)
                        {
                            var user = org.Retrieve(userRef.LogicalName, userRef.Id, new ColumnSet("domainname"));
                            var username = user.GetAttributeValue<string>("domainname");
                            var userTeamsDetails = (User)table.Execute(TableOperation.Retrieve<User>(username, "")).Result;

                            if (userTeamsDetails == null)
                                continue;

                            MicrosoftAppCredentials.TrustServiceUrl(userTeamsDetails.ServiceUrl);
                            var client = new ConnectorClient(new Uri(userTeamsDetails.ServiceUrl), _config.GetValue<string>("MicrosoftAppId"), _config.GetValue<string>("MicrosoftAppPassword"));

                            if (avatarUrl == null)
                            {
                                var sender = org.Retrieve("systemuser", postComment.GetAttributeValue<EntityReference>("createdby").Id, new ColumnSet("domainname"));
                                avatarUrl = await GetAvatarUrlAsync(userTeamsDetails.TenantId, sender.GetAttributeValue<string>("domainname"), postComment.GetAttributeValue<EntityReference>("createdby").Name);
                            }

                            // Create or get existing chat conversation with user
                            var parameters = new ConversationParameters
                            {
                                Bot = new ChannelAccount("28:" + _config.GetValue<string>("MicrosoftAppId")),
                                Members = new[] { new ChannelAccount(userTeamsDetails.UserId) },
                                ChannelData = new TeamsChannelData
                                {
                                    Tenant = new TenantInfo(userTeamsDetails.TenantId),
                                },
                            };

                            var response = await client.Conversations.CreateConversationAsync(parameters);

                            // Construct the message to post to conversation
                            var model = new PostNotify
                            {
                                Regarding = entityRef,
                                Post = post,
                                Comment = postComment == post ? null : postComment,
                                Links = GetChain(postComment.Id, userRef.Id, entityRelationships)
                            };

                            EnsureChainDetails(org, model.Links);

                            var newActivity = new Activity
                            {
                                //Text = model.ToString(),
                                Type = ActivityTypes.Message,
                                Conversation = new ConversationAccount
                                {
                                    Id = response.Id
                                },
                                Attachments = new List<Attachment>
                                {
                                    ToAdaptiveCard(model, avatarUrl)
                                }
                            };

                            // Post the message to chat conversation with user
                            await client.Conversations.SendToConversationAsync(response.Id, new Activity { Type = ActivityTypes.Typing });
                            await Task.Delay(TimeSpan.FromSeconds(2));
                            await client.Conversations.SendToConversationAsync(response.Id, newActivity);

                            // Update the user to record the details of which post should be replied to if the user sends a message back
                            var updatedUser = new User(username)
                            {
                                LastDomainName = DomainName,
                                LastPostId = post.Id,
                                ETag = userTeamsDetails.ETag
                            };
                            var update = TableOperation.Merge(updatedUser);

                            try
                            {
                                await table.ExecuteAsync(update);
                            }
                            catch (StorageException ex)
                            {
                                if (ex.RequestInformation.HttpStatusCode != 412)
                                    throw;
                            }
                        }
                    }
                }

                return Ok();
            }
            catch (Exception ex)
            {
                return Problem(ex.ToString(), statusCode: 500);
            }
        }

        private async Task<string> GetAvatarUrlAsync(string tenantId, string upn, string name)
        {
            var token = await GetApplicationTokenAsync(tenantId);

            var req = WebRequest.CreateHttp($"https://graph.microsoft.com/v1.0/users/{upn}/photos/48x48/$value");
            req.Headers[HttpRequestHeader.Authorization] = "Bearer " + token;
            req.Headers[HttpRequestHeader.ContentType] = "image/jpg";

            try
            {
                using (var resp = req.GetResponse())
                using (var stream = resp.GetResponseStream())
                {
                    var bytes = new byte[resp.ContentLength];
                    stream.Read(bytes, 0, bytes.Length);
                    return "data:" + resp.ContentType + ";base64," + Convert.ToBase64String(bytes);
                }
            }
            catch (WebException)
            {
                using (var bitmap = new Bitmap(48, 48))
                using (var g = Graphics.FromImage(bitmap))
                {
                    g.FillRectangle(Brushes.Green, 0, 0, 48, 48);

                    var words = name.Split(' ');
                    var text = words[0][0].ToString();

                    if (words.Length > 1)
                        text += words.Last()[0];

                    text = text.ToUpper();

                    using (var font = new Font("Arial", 12, FontStyle.Bold))
                    {
                        var size = g.MeasureString(text, font);
                        g.DrawString(text, font, Brushes.White, (bitmap.Width - size.Width) / 2, (bitmap.Height - size.Height) / 2);
                    }

                    using (var stream = new MemoryStream())
                    {
                        bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Jpeg);
                        return "data:image/jpeg;base64," + Convert.ToBase64String(stream.ToArray());
                    }
                }
            }
        }

        private async Task<string> GetApplicationTokenAsync(string tenantId)
        {
            var cca = ConfidentialClientApplicationBuilder.Create(_config.GetValue<string>("MicrosoftAppId"))
                .WithTenantId(tenantId)
                .WithRedirectUri("msal" + _config.GetValue<string>("MicrosoftAppId") + "://auth")
                .WithClientSecret(_config.GetValue<string>("MicrosoftAppPassword"))
                .Build();

            var result = await cca.AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();
            return result.AccessToken;
        }

        private void EnsureChainDetails(IOrganizationService org, Link[] links)
        {
            foreach (var link in links)
            {
                if (!String.IsNullOrEmpty(link.From.Name) ||
                    link.From.LogicalName == "post" ||
                    link.From.LogicalName == "postcomment")
                    continue;

                var nameAttr = GetEntityMetadata(org, link.From.LogicalName).PrimaryNameAttribute;
                var entity = org.Retrieve(link.From.LogicalName, link.From.Id, new ColumnSet(nameAttr));
                link.From.Name = entity.GetAttributeValue<string>(nameAttr);
            }
        }

        private EntityMetadata GetEntityMetadata(IOrganizationService org, string logicalName)
        {
            var instanceCache = _metadata.GetOrAdd(DomainName, _ => new ConcurrentDictionary<string, EntityMetadata>());
            return instanceCache.GetOrAdd(logicalName, ln =>
            {
                var req = new RetrieveEntityRequest
                {
                    LogicalName = ln,
                    EntityFilters = EntityFilters.Entity | EntityFilters.Attributes
                };
                var resp = (RetrieveEntityResponse)org.Execute(req);
                return resp.EntityMetadata;
            });
        }

        private Link[] GetChain(Guid from, Guid to, Dictionary<Guid, Link> entityRelationships)
        {
            var chain = new List<Link>();

            GetChain(chain, from, to, entityRelationships);

            return chain.ToArray();
        }

        private void GetChain(List<Link> chain, Guid from, Guid to, Dictionary<Guid, Link> entityRelationships)
        {
            var link = entityRelationships[to];
            chain.Insert(0, link);

            if (link.From.Id != from)
                GetChain(chain, from, link.From.Id, entityRelationships);
        }

        private static Entity GetFullPostText(IOrganizationService org, EntityReference entityRef, Entity post,
            ref Entity postComment)
        {
            // Retrieve the full wall for this record to expand out any standard posts
            var wallPage = 1;

            while (true)
            {
                var wall = (RetrieveRecordWallResponse)org.Execute(new RetrieveRecordWallRequest
                {
                    Entity = entityRef,
                    CommentsPerPost = 10,
                    PageSize = 10,
                    PageNumber = wallPage,
                    Source = post.GetAttributeValue<OptionSetValue>("source")
                });

                var foundPost = false;

                foreach (var wallPost in wall.EntityCollection.Entities)
                {
                    if (wallPost.Id == post.Id)
                    {
                        if (post == postComment)
                            postComment = wallPost;

                        post = wallPost;
                        foundPost = true;
                        break;
                    }
                }

                if (foundPost || !wall.EntityCollection.MoreRecords)
                    break;

                wallPage++;
            }

            if (postComment == post)
            {
                postComment = post;
            }
            else
            {
                EntityCollection comments;
                if (post.RelatedEntities.TryGetValue(new Relationship("Post_Comments"), out comments))
                {
                    foreach (var wallComment in comments.Entities)
                    {
                        if (wallComment.Id == postComment.Id)
                        {
                            postComment = wallComment;
                            break;
                        }
                    }
                }
            }
            return post;
        }

        private void AddUserFollows(HashSet<EntityReference> usersToNotify, IOrganizationService org, Dictionary<Guid, Link> entityRelationships)
        {
            foreach (var user in usersToNotify.ToList())
            {
                var followsQry = new QueryByAttribute("postfollow");
                followsQry.AddAttributeValue("regardingobjectid", user.Id);
                followsQry.ColumnSet = new ColumnSet("ownerid");

                foreach (var follow in org.RetrieveMultiple(followsQry).Entities)
                {
                    var userId = follow.GetAttributeValue<EntityReference>("ownerid");

                    if (usersToNotify.Add(userId))
                    {
                        entityRelationships.Add(userId.Id, new Link { From = user, Description = "Followed By" });
                    }
                }
            }
        }

        private void ExpandTeamsToUsers(HashSet<EntityReference> usersToNotify, IOrganizationService org, Dictionary<Guid, Link> entityRelationships)
        {
            foreach (var team in usersToNotify.Where(usr => usr.LogicalName == "team").ToList())
            {
                var usersQry = new QueryByAttribute("connection");
                usersQry.AddAttributeValue("record1id", team.Id);
                usersQry.AddAttributeValue("record1roleid", new Guid("8F443BC5-19E3-E611-80C8-00155D007101"));
                usersQry.AddAttributeValue("record2roleid", new Guid("194FDF45-1AE3-E611-80C8-00155D007101"));
                usersQry.ColumnSet = new ColumnSet("record2id");

                foreach (var connection in org.RetrieveMultiple(usersQry).Entities)
                {
                    var userId = connection.GetAttributeValue<EntityReference>("record2id");
                    if (usersToNotify.Add(userId))
                        entityRelationships.Add(userId.Id, new Link { From = team, Description = "Connected To" });
                }
            }
        }

        private void GetInterestedUsers(Entity post, Entity comment, Entity entity, IOrganizationService org, HashSet<EntityReference> userIds, Dictionary<Guid, Link> entityRelationships)
        {
            var processedEntities = new HashSet<Guid>();

            // Get any users mentioned in the post
            var userRegex = new Regex(@"@\[(?<etc>[0-9]+),(?<id>[a-z0-9-]+),", RegexOptions.IgnoreCase);

            foreach (Match user in userRegex.Matches(comment.GetAttributeValue<string>("text")))
                FollowMention(Int32.Parse(user.Groups["etc"].Value), new Guid(user.Groups["id"].Value), org, processedEntities, new Link { From = comment.ToEntityReference(), Description = "Mentions" }, userIds, entityRelationships);

            // Get the user who posted the comment
            var createdBy = comment.GetAttributeValue<EntityReference>("createdby");
            if (userIds.Add(createdBy))
                entityRelationships.Add(createdBy.Id, new Link { From = comment.ToEntityReference(), Description = "Posted By" });

            // Get the user who posted the original post
            createdBy = post.GetAttributeValue<EntityReference>("createdby");
            if (userIds.Add(createdBy))
                entityRelationships.Add(createdBy.Id, new Link { From = post.ToEntityReference(), Description = "Posted By" });

            // Add any user who has also replied to this same post
            var replyQry = new QueryByAttribute("postcomment");
            replyQry.AddAttributeValue("postid", post.Id);
            replyQry.ColumnSet = new ColumnSet("createdby");

            foreach (var reply in org.RetrieveMultiple(replyQry).Entities)
            {
                createdBy = reply.GetAttributeValue<EntityReference>("createdby");
                if (userIds.Add(createdBy))
                    entityRelationships.Add(createdBy.Id, new Link { From = entity.ToEntityReference(), Description = "Also Commented On By" });
            }

            // Get anyone otherwise interested in the record being posted on
            GetInterestedUsers(entity, org, processedEntities, userIds, entityRelationships);
        }

        private void FollowMention(int etc, Guid id, IOrganizationService org, HashSet<Guid> processedEntities, Link relationship, HashSet<EntityReference> userIds, Dictionary<Guid, Link> entityRelationships)
        {
            if (processedEntities.Contains(id))
                return;

            if (etc == 8)
            {
                // User
                var userRef = new EntityReference("systemuser", id);
                if (userIds.Add(userRef))
                    entityRelationships.Add(id, relationship);
            }
            else if (etc == 9)
            {
                // Team
                var teamRef = new EntityReference("team", id);
                if (userIds.Add(teamRef))
                    entityRelationships.Add(id, relationship);
            }
            else
            {
                // Other record
                // Get the logical name from the type code
                // TODO: Make this more efficient with a cache and using RetrieveMetadataChangesRequest with appropriate filter
                var metadataReq = new RetrieveAllEntitiesRequest();
                metadataReq.EntityFilters = EntityFilters.Entity;
                var metadataResp = (RetrieveAllEntitiesResponse)org.Execute(metadataReq);
                var metadata = metadataResp.EntityMetadata.SingleOrDefault(e => e.ObjectTypeCode == etc);

                if (metadata != null)
                {
                    var entity = org.Retrieve(metadata.LogicalName, id, new ColumnSet(true));

                    if (!entityRelationships.ContainsKey(id))
                        entityRelationships.Add(id, relationship);

                    GetInterestedUsers(entity, org, processedEntities, userIds, entityRelationships);
                }
            }
        }

        private void GetInterestedUsers(Entity entity, IOrganizationService org, HashSet<Guid> processedEntities, HashSet<EntityReference> userIds, Dictionary<Guid, Link> entityRelationships)
        {
            if (!processedEntities.Add(entity.Id))
                return;

            // Add any user explicitly linked to from the parent record 
            // Include owner, exclude createdby and modifiedby as they can set to workflow owners who aren't interested
            // in specific records
            foreach (var attribute in entity.Attributes)
            {
                if (attribute.Key == "createdby" || attribute.Key == "modifiedby")
                    continue;

                var userRef = attribute.Value as EntityReference;

                if (userRef != null && (userRef.LogicalName == "systemuser" || userRef.LogicalName == "team"))
                {
                    if (userIds.Add(userRef))
                        entityRelationships.Add(userRef.Id, new Link { From = entity.ToEntityReference(), Description = GetEntityMetadata(org, entity.LogicalName).Attributes.Single(a => a.LogicalName == attribute.Key).DisplayName.UserLocalizedLabel.Label });
                }
            }

            // Add any user who follows the parent record.
            var entityFollowsQry = new QueryByAttribute("postfollow");
            entityFollowsQry.AddAttributeValue("regardingobjectid", entity.Id);
            entityFollowsQry.ColumnSet = new ColumnSet("ownerid");

            foreach (var entityFollow in org.RetrieveMultiple(entityFollowsQry).Entities)
            {
                var userRef = entityFollow.GetAttributeValue<EntityReference>("ownerid");

                if (userIds.Add(userRef))
                    entityRelationships.Add(userRef.Id, new Link { From = entity.ToEntityReference(), Description = "Followed By" });
            }

            // Recurse into any accounts linked to this record
            var accountIds = new HashSet<Guid>();

            foreach (var attribute in entity.Attributes)
            {
                var accountRef = attribute.Value as EntityReference;

                if (accountRef != null && accountRef.LogicalName == "account" && !processedEntities.Contains(accountRef.Id))
                {
                    accountIds.Add(accountRef.Id);

                    if (!entityRelationships.ContainsKey(accountRef.Id))
                        entityRelationships.Add(accountRef.Id, new Link { From = entity.ToEntityReference(), Description = GetEntityMetadata(org, entity.LogicalName).Attributes.Single(a => a.LogicalName == attribute.Key).DisplayName.UserLocalizedLabel.Label });
                }
            }

            foreach (var accountId in accountIds)
            {
                var account = org.Retrieve("account", accountId, new ColumnSet("ownerid"));

                // Recurse into the account
                GetInterestedUsers(account, org, processedEntities, userIds, entityRelationships);
            }
        }

        private Attachment ToAdaptiveCard(PostNotify post, string avatarUrl)
        {
            var card = new AdaptiveCard("1.2")
            {
                Body = new List<AdaptiveElement>
                {
                    // Header
                    new AdaptiveTextBlock
                    {
                        Size = AdaptiveTextSize.Large,
                        Weight = AdaptiveTextWeight.Bolder,
                        Text = GetLink(DomainName, post.Regarding)
                    },

                    // Author
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Style = AdaptiveImageStyle.Person,
                                        Size = AdaptiveImageSize.Small,
                                        Url = new Uri(avatarUrl)
                                    }
                                },
                                Width = "auto",
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Text = GetLink(DomainName, (EntityReference) (post.Comment ?? post.Post)["createdby"]),
                                        Wrap = true
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = $"{{{{DATE({(post.Comment ?? post.Post)["createdon"]:yyyy-MM-ddTHH:mm:ssZ})}}}} {{{{TIME({(post.Comment ?? post.Post)["createdon"]:yyyy-MM-ddTHH:mm:ssZ})}}}}",
                                        Size = AdaptiveTextSize.Small,
                                        IsSubtle = true,
                                        Spacing = AdaptiveSpacing.None
                                    }
                                },
                                Width = "stretch"
                            }
                        }
                    },

                    // Post
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = FormatPostText((string) post.Post["text"]),
                                        IsSubtle = post.Comment != null
                                    }
                                },
                                Style = AdaptiveContainerStyle.Emphasis
                            }
                        }
                    }
                },
                Actions = new List<AdaptiveAction>
                {
                    // Reply
                    new AdaptiveShowCardAction
                    {
                        Title = "Reply",
                        Card = new AdaptiveCard("1.2")
                        {
                            Body = new List<AdaptiveElement>
                            {
                                new AdaptiveTextInput
                                {
                                    Id = "comment",
                                    Placeholder = "Add a comment",
                                    IsMultiline = true
                                }
                            },
                            Actions = new List<AdaptiveAction>
                            {
                                new AdaptiveSubmitAction
                                {
                                    Title = "OK",
                                    Data = new
                                    {
                                        DomainName,
                                        PostId = post.Post.Id
                                    }
                                }
                            }
                        }
                    },

                    // Chain
                    new AdaptiveShowCardAction
                    {
                        Title = "Why did I get this?",
                        Card = new AdaptiveCard("1.2")
                        {
                            Body = post.Links
                                .SelectMany(link => new AdaptiveElement[]
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = link.From.Id == post.Comment?.Id ? "*This comment*" :
                                                link.From.Id == post.Post.Id ? "*This post*" :
                                                GetLink(DomainName, link.From),
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                        Spacing = AdaptiveSpacing.Small
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = "🔻",
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                        Spacing = AdaptiveSpacing.Small,
                                        Size = AdaptiveTextSize.Small
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = link.Description,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                        IsSubtle = true,
                                        Spacing = AdaptiveSpacing.Small,
                                        Size = AdaptiveTextSize.Small
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = "🔻",
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                        Spacing = AdaptiveSpacing.Small,
                                        Size = AdaptiveTextSize.Small
                                    },
                                })
                                .Concat(new AdaptiveElement[]
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = "You",
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                                        Spacing = AdaptiveSpacing.Small
                                    }
                                })
                                .ToList()
                        }
                    },

                    // Open Record
                    new AdaptiveOpenUrlAction
                    {
                        Title = "View Record",
                        Url = new Uri(GetUrl(DomainName, post.Regarding))
                    }
                }
            };

            if (post.Comment != null)
            {
                ((AdaptiveColumnSet)card.Body.Last()).Columns[0].Items.Add(new AdaptiveTextBlock
                {
                    Text = $"by {GetLink(DomainName, (EntityReference)post.Post["createdby"])} at {{{{DATE({post.Post["createdon"]:yyyy-MM-ddTHH:mm:ssZ})}}}} {{{{TIME({post.Post["createdon"]:yyyy-MM-ddTHH:mm:ssZ})}}}}",
                    Size = AdaptiveTextSize.Small
                });

                card.Body.Add(new AdaptiveColumnSet
                {
                    Columns = new List<AdaptiveColumn>
                    {
                        new AdaptiveColumn
                        {
                            Width = "20px"
                        },
                        new AdaptiveColumn
                        {
                            Width = "stretch",
                            Items = new List<AdaptiveElement>
                            {
                                new AdaptiveTextBlock
                                {
                                    Text = FormatPostText((string)post.Comment["text"]),
                                    Wrap = true
                                }
                            },
                            Spacing = AdaptiveSpacing.Medium,
                            Style = AdaptiveContainerStyle.Accent
                        }
                    }
                });
            }

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card
            };
        }

        private string GetUrl(string domainName, EntityReference entityRef)
        {
            return $"https://{domainName}/main.aspx?etn={entityRef.LogicalName}&pagetype=entityrecord&id={entityRef.Id}";
        }

        private string GetLink(string domainName, EntityReference entityRef)
        {
            return $"[{entityRef.Name}]({GetUrl(domainName, entityRef)})";
        }

        private string FormatPostText(string text)
        {
            // Tags in text stored as @[otc,guid,"text"]
            text = Regex.Replace(
                text,
                @"@\[(?<etc>[0-9]+),(?<id>[-0-9a-zA-Z]+),""(?<name>[^""]*)""]",
                mention => $"[{mention.Groups["name"].Value}](https://{DomainName}/main.aspx?etc={mention.Groups["etc"].Value}&id={mention.Groups["id"].Value}&pagetype=entityrecord)");

            return text.Replace("\n", "\n\n");
        }
    }

    public class Link
    {
        public EntityReference From { get; set; }
        public string Description { get; set; }
    }

    public class PostNotify
    {
        public EntityReference Regarding { get; set; }
        public Entity Post { get; set; }
        public Entity Comment { get; set; }
        public Link[] Links { get; set; }
    }
}
