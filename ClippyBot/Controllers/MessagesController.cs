using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using Microsoft.Bot.Connector;
using Newtonsoft.Json;
using Microsoft.Bot.Builder.Dialogs;
using ClippyBot.Dialogs;

namespace ClippyBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {
            if (activity.Type == ActivityTypes.Event)
            {
                if (activity.Name == "initialize")
                {
                    // Get the Office host from activity Properties then save it into BotState
                    var host = activity.Value.ToString();
                    var state = activity.GetStateClient();
                    var userdata = state.BotState.GetUserData(activity.ChannelId, activity.From.Id);
                    userdata.SetProperty<string>("host", host);
                    userdata.SetProperty<string>("user", activity.From.Name);
                    state.BotState.SetUserData(activity.ChannelId, activity.From.Id, userdata);

                    // Route the activity to the correct dialog
                    await routeActivity(activity);
                }
                else if (activity.Name == "confirmation")
                {
                    // Respond with the results in the bot
                    ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));
                    Activity reply = activity.CreateReply("Looks like the Clippy Bot was successful...let's try that again.");
                    await connector.Conversations.ReplyToActivityAsync(reply);
                    await routeActivity(activity);
                }
            }
            else if (activity.Type == ActivityTypes.Message)
            {
                // Route the activity to the correct dialog
                await routeActivity(activity);
            }
            else
            {
                HandleSystemMessage(activity);
            }
            var response = Request.CreateResponse(HttpStatusCode.OK);
            return response;
        }

        private async Task routeActivity(Activity activity)
        {
            // Make sure we know the host
            var state = activity.GetStateClient();
            var userdata = state.BotState.GetUserData(activity.ChannelId, activity.From.Id);
            var host = userdata.GetProperty<string>("host");
            switch (host)
            {
                case "Word":
                    await Conversation.SendAsync(activity, () => new WordDialog());
                    break;
                case "Excel":
                    await Conversation.SendAsync(activity, () => new ExcelDialog());
                    break;
                case "Outlook":
                    await Conversation.SendAsync(activity, () => new OutlookDialog());
                    break;
                default:
                    ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));
                    Activity reply = activity.CreateReply($"Sorry, I can't figure out where you are running me from. You may not have given me enough time to initialize.");
                    await connector.Conversations.ReplyToActivityAsync(reply);
                    break;
            }
        }

        private Activity HandleSystemMessage(Activity message)
        {
            if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing tha the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {
            }

            return null;
        }
    }
}