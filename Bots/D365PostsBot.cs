using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Azure;
using Azure.Core;
using Azure.Data.Tables;
using MarkMpn.D365PostsBot.Models;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Entity = Microsoft.Xrm.Sdk.Entity;

namespace MarkMpn.D365PostsBot.Bots
{
    public class D365PostsBot : TeamsActivityHandler
    {
        private readonly IConfiguration _config;
        private readonly TokenCredential _credential;

        public D365PostsBot(
            IConfiguration config,
            TokenCredential credential)
        {
            _config = config;
            _credential = credential;
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
                    var table = new TableClient(new Uri(connectionString), "users", _credential);

                    User user = null;

                    try
                    {
                        user = (await table.GetEntityAsync<User>(username, "", cancellationToken: cancellationToken)).Value;
                    }
                    catch (RequestFailedException ex) when (ex.Status == 404)
                    {
                    }

                    if (user == null)
                    {
                        try
                        {
                            user = (await table.GetEntityAsync<User>(username.ToLowerInvariant(), "", cancellationToken: cancellationToken)).Value;
                        }
                        catch (RequestFailedException ex) when (ex.Status == 404)
                        {
                        }
                    }

                    if (user == null)
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text("Sorry, I couldn't find your user details. Please remove and re-add the D365 Posts Bot app in Teams and try again"), cancellationToken: cancellationToken);
                        return;
                    }

                    if (user.LastPostId == null)
                    {
                        await turnContext.SendActivityAsync(MessageFactory.Text("Sorry, I couldn't find a message to reply to. Please use the Reply button on the post you want to reply to."), cancellationToken: cancellationToken);
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

                await turnContext.SendActivityAsync(new Activity { Type = ActivityTypes.Typing }, cancellationToken: cancellationToken);

                try
                {
                    using (var org = new ServiceClient(new Uri("https://" + domainName), _config.GetValue<string>("MicrosoftAppId"), _config.GetValue<string>("MicrosoftAppPassword"), true, null))
                    {
                        // Find the Dataverse user details
                        var userQry = new QueryByAttribute("systemuser") { ColumnSet = new ColumnSet("systemuserid") };
                        userQry.AddAttributeValue("domainname", username);
                        var users = await org.RetrieveMultipleAsync(userQry);

                        if (users.Entities.Count == 0)
                            throw new ApplicationException("Could not find your user account in D365");

                        org.CallerId = users.Entities[0].Id;

                        var postComment = new Entity("postcomment")
                        {
                            ["postid"] = new EntityReference("post", postId),
                            ["text"] = message
                        };

                        await org.CreateAsync(postComment);
                    }

                    await turnContext.SendActivityAsync(MessageFactory.Text("Your reply has been sent"), cancellationToken: cancellationToken);
                }
                catch (Exception ex)
                {
                    await turnContext.SendActivityAsync(MessageFactory.Text("There was an error sending your reply back to D365: " + ex.Message), cancellationToken: cancellationToken);
                }
            }
            catch (Exception ex)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(ex.ToString()), cancellationToken: cancellationToken);
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
                var table = new TableClient(new Uri(connectionString), "users", _credential);
                await table.CreateIfNotExistsAsync();

                foreach (var member in membersAdded)
                {
                    if (member.Id != turnContext.Activity.Recipient.Id)
                    {
                        var userId = member.Id;
                        var username = ((TeamsChannelAccount)member).UserPrincipalName.ToLowerInvariant();

                        // Store details
                        var user = new User(username)
                        {
                            UserId = userId,
                            TenantId = tenantId,
                            ServiceUrl = serviceUrl
                        };


                        try
                        {
                            await table.AddEntityAsync(user, cancellationToken);

                            await turnContext.SendActivityAsync(MessageFactory.Text("Welcome!"), cancellationToken: cancellationToken);
                        }
                        catch (RequestFailedException ex) when (ex.Status == 409)
                        {
                            // Don't throw errors if we've seen this user before
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("Error: " + ex.ToString()), cancellationToken: cancellationToken);
            }
        }
    }
}
