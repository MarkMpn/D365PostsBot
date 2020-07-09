using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using MarkMpn.D365PostsBot.Models;
using Microsoft.Azure.Cosmos.Table;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Microsoft.PowerPlatform.Cds.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Entity = Microsoft.Xrm.Sdk.Entity;

namespace MarkMpn.D365PostsBot.Bots
{
    public class D365PostsBot : TeamsActivityHandler
    {
        private readonly IConfiguration _config;

        public D365PostsBot(IConfiguration config)
        {
            _config = config;
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var member = await TeamsInfo.GetMemberAsync(turnContext, turnContext.Activity.From.Id, cancellationToken);
                var username = member.UserPrincipalName;

                string message;
                Guid postId;
                string domainName;

                if (!String.IsNullOrEmpty(turnContext.Activity.Text))
                {
                    var connectionString = _config.GetConnectionString("Storage");
                    var storageAccount = CloudStorageAccount.Parse(connectionString);
                    var tableClient = storageAccount.CreateCloudTableClient();
                    var table = tableClient.GetTableReference("users");
                    var user = (User)table.Execute(TableOperation.Retrieve<User>(username, "")).Result;

                    if (user == null)
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text("Sorry, I couldn't find your user details. Please remove and re-add the D365 Posts Bot app in Teams and try again"));
                        return;
                    }

                    if (user.LastPostId == null)
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text("Sorry, I couldn't find a message to reply to. Please use the Reply button on the post you want to reply to."));
                        return;
                    }

                    message = turnContext.Activity.Text;
                    postId = user.LastPostId.Value;
                    domainName = user.LastDomainName;
                }
                else
                {
                    dynamic val = turnContext.Activity.Value;

                    message = val.comment;
                    postId = val.PostId;
                    domainName = val.DomainName;
                }

                await turnContext.SendActivityAsync(new Activity { Type = ActivityTypes.Typing });

                try
                {
                    using (var org = new CdsServiceClient(new Uri("https://" + domainName), _config.GetValue<string>("MicrosoftAppId"), _config.GetValue<string>("MicrosoftAppPassword"), true, null))
                    {
                        // Find the CDS user details
                        var userQry = new QueryByAttribute("systemuser") { ColumnSet = new ColumnSet("systemuserid") };
                        userQry.AddAttributeValue("domainname", username);
                        var users = org.RetrieveMultiple(userQry);

                        if (users.Entities.Count == 0)
                            throw new ApplicationException("Could not find your user account in D365");

                        org.CallerId = users.Entities[0].Id;

                        var postComment = new Entity("postcomment")
                        {
                            ["postid"] = new EntityReference("post", postId),
                            ["text"] = message
                        };

                        org.Create(postComment);
                    }

                    await turnContext.SendActivityAsync(MessageFactory.Text("Your reply has been sent"));
                }
                catch (Exception ex)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text("There was an error sending your reply back to D365: " + ex.Message));
                }
            }
            catch (Exception ex)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(ex.ToString()));
            }
        }

        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                var teamConversationData = turnContext.Activity.GetChannelData<TeamsChannelData>();
                var tenantId = teamConversationData.Tenant.Id;
                var serviceUrl = turnContext.Activity.ServiceUrl;

                var connectionString = _config.GetConnectionString("Storage");
                var storageAccount = CloudStorageAccount.Parse(connectionString);
                var tableClient = storageAccount.CreateCloudTableClient();
                var table = tableClient.GetTableReference("users");
                table.CreateIfNotExists();

                foreach (var member in membersAdded)
                {
                    if (member.Id != turnContext.Activity.Recipient.Id)
                    {
                        var userId = member.Id;
                        var username = ((TeamsChannelAccount)member).UserPrincipalName;

                        // Store details
                        var user = new User(username)
                        {
                            UserId = userId,
                            TenantId = tenantId,
                            ServiceUrl = serviceUrl
                        };

                        var insert = TableOperation.Insert(user);

                        try
                        {
                            await table.ExecuteAsync(insert);

                            await turnContext.SendActivityAsync(MessageFactory.Text("Welcome!"));
                        }
                        catch (StorageException ex)
                        {
                            // Don't throw errors if we've seen this user before
                            if (ex.RequestInformation.HttpStatusCode != 409)
                                throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("Error: " + ex.ToString()));
            }
        }
    }
}
