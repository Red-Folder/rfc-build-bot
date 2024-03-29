#r "Newtonsoft.Json"
#load "EchoDialog.csx"

using System;
using System.Net;
using System.Threading;
using Newtonsoft.Json;
using System.IO;

using Microsoft.Bot.Builder.Azure;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;

public static async Task<object> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info($"Webhook was triggered!");

    // Initialize the azure bot
    using (BotService.Initialize())
    {
        // Deserialize the incoming activity
        string jsonContent = await req.Content.ReadAsStringAsync();
        var activity = JsonConvert.DeserializeObject<Activity>(jsonContent);
        
        // authenticate incoming request and add activity.ServiceUrl to MicrosoftAppCredentials.TrustedHostNames
        // if request is authenticated
        if (!await BotService.Authenticator.TryAuthenticateAsync(req, new [] {activity}, CancellationToken.None))
        {
            return BotAuthenticator.GenerateUnauthorizedResponse(req);
        }
        
        if (activity != null)
        {
            // Store Conversation for later use
            StoreConversation(activity, log); 
            
            // one of these will have an interface and process it
            switch (activity.GetActivityType())
            {
                case ActivityTypes.Message:
                    await Conversation.SendAsync(activity, () => new EchoDialog());
                    break;
                case ActivityTypes.ConversationUpdate:
                    var client = new ConnectorClient(new Uri(activity.ServiceUrl));
                    IConversationUpdateActivity update = activity;
                    if (update.MembersAdded.Any())
                    {
                        var reply = activity.CreateReply();
                        var newMembers = update.MembersAdded?.Where(t => t.Id != activity.Recipient.Id);
                        foreach (var newMember in newMembers)
                        {
                            reply.Text = "Welcome";
                            if (!string.IsNullOrEmpty(newMember.Name))
                            {
                                reply.Text += $" {newMember.Name}";
                            }
                            reply.Text += "!";
                            await client.Conversations.ReplyToActivityAsync(reply);
                        }
                    }

                    if (update.MembersRemoved.Where(t => t.Id == activity.Recipient.Id).Any())
                    {
                        RemoveConversation(activity, log);
                    }
                    break;
                case ActivityTypes.ContactRelationUpdate:
                case ActivityTypes.Typing:
                case ActivityTypes.DeleteUserData:
                case ActivityTypes.Ping:
                default:
                    log.Error($"Unknown activity type ignored: {activity.GetActivityType()}");
                    break;
            }
        }
        return req.CreateResponse(HttpStatusCode.Accepted);
    }    
}

private static void StoreConversation(Activity activity, TraceWriter log)
{
    var homeFolder = Environment.GetEnvironmentVariable("HOME");
    var storageFolder = homeFolder + "\\data\\" + CleanFolderName(activity.Conversation.Id, log);
    if (!Directory.Exists(storageFolder))
    {
        Directory.CreateDirectory(storageFolder);
        File.WriteAllText(storageFolder + "\\serviceUrl.txt", activity.ServiceUrl);
        File.WriteAllText(storageFolder + "\\user.json", JsonConvert.SerializeObject(activity.From));
        File.WriteAllText(storageFolder + "\\bot.json", JsonConvert.SerializeObject(activity.Recipient));
        File.WriteAllText(storageFolder + "\\conversation.json", JsonConvert.SerializeObject(activity.Conversation));
    }
}

private static void RemoveConversation(Activity activity, TraceWriter log)
{
    var homeFolder = Environment.GetEnvironmentVariable("HOME");
    var storageFolder = homeFolder + "\\data\\" + CleanFolderName(activity.Conversation.Id, log);
    if (Directory.Exists(storageFolder))
    {
        Directory.Delete(storageFolder, true);
    }
}

private static string CleanFolderName(string foldername, TraceWriter log)
{
    log.Info("Cleaning filename");
    log.Info("Before: " + foldername);
    var invalids = System.IO.Path.GetInvalidFileNameChars();
    var newFolderName = String.Join("_", foldername.Split(invalids, StringSplitOptions.RemoveEmptyEntries) ).TrimEnd('.');
    log.Info("After: " + newFolderName);
    return newFolderName;
}
