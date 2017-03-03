using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace ClippyBot.Dialogs
{
    [Serializable]
    public class OutlookDialog : IDialog<IMessageActivity>
    {
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var msg = await result;

            string[] options = new string[] { "Reply", "Reply All" };
            var user = context.UserData.Get<string>("user");
            string prompt = $"Hey {user}, I see you are running the Clippy Bot in Outlook. For this mail, I can help you with the following:";
            PromptDialog.Choice(context, async (IDialogContext choiceContext, IAwaitable<string> choiceResult) =>
            {
                var selection = await choiceResult;
                OfficeOperation op = (OfficeOperation)Enum.Parse(typeof(OfficeOperation), selection.Replace(" ", ""));

                // Send the operation through the backchannel using Event activity
                var reply = choiceContext.MakeMessage() as IEventActivity;
                reply.Type = "event";
                reply.Name = "officeOperation";
                reply.Value = op.ToString();
                await choiceContext.PostAsync((IMessageActivity)reply);
            }, options, prompt);
        }
    }
}