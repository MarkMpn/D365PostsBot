# 📢 D365 Posts Bot

D365 Posts Bot sends notifications to users via Teams when a post is added to D365 that is related to them. The user can reply to posts
directly from Teams and the bot will save the reply to D365.

The [free hosted version of the bot](https://bot.markcarrington.dev/) can be used quickly, or you can self-host it.

## Self-Hosting

### 1. Creating the bot and Teams app

In order to self-host the bot, you must first create your own Web App Bot resource in Azure. This will provide you with an App Service
to host the bot in, and all the registration details for the bot itself. You also need to enable the Teams channel for your bot and
a corresponding Teams app. You can't reuse the existing one I've created for the hosted version as your bot will have a different ID.

[Detailed instructions for this step are available on my blog](https://markcarrington.dev/2020/06/02/creating-a-bot-pt-1-getting-started/).

### 2. Creating a storage account in Azure

You also need to create a storage account to hold some basic details of the users of your bot. You can create this using the
[Azure portal](https://portal.azure.com) - create a standard storage account and note the connection string.

### 3. Update the configuration file

In the `appsettings.json` file, add the following settings:

```json
{
  "MicrosoftAppId": "guid",
  "MicrosoftAppPassword": "key",
  "ConnectionStrings": {
    "Storage": "connectionstring"
  },
  "WebhookKey": "key"
}
```

The `MicrosoftAppId` should be set to the ID of the bot resource you created in step 1, and the `MicrosoftAppPassword` should be
the associated client secret.

The `ConnectionStrings.Storage` entry should be set to the connection string for the storage account you created in step 2.

The `WebhookKey` entry can be set to any value you like. Whatever value you use here you'll need to use again later on, so keep
a note of it.

### 4. Register the web hook in D365

Using the Plugin Registration Tool, add a new web hook. The endpoint URL should be in the format `https://your.domain.com/api/notification`,
set the Authentication type to `WebhookKey` and the Value to the key you entered into the configuration file in the previous step.

Once the web hook is available, add steps to trigger it when `post` and `postcomment` records are created.

[Details of this process are available on my blog](https://markcarrington.dev/2020/06/15/creating-a-bot-pt-5-getting-notifications-from-d365/)

### 5. Install the Teams app

You need to deploy the Teams app you created in step 1 to each user. Each user can do this manually or you can use a policy to
deploy it automatically.
[Details on both methods are available on the website for the hosted version](https://bot.markcarrington.dev/teamsapp.html),
but use your own custom Teams app instead of downloading the one referenced in those instructions.
