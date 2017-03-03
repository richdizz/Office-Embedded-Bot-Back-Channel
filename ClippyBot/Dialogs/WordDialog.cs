using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace ClippyBot.Dialogs
{
    [Serializable]
    public class WordDialog : IDialog<IMessageActivity>
    {
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var msg = await result;

            string[] options = new string[] { "Image", "Paragraph" };
            string prompt = $"I see you are running the Clippy Bot in Word. Which of the following would you like me to insert into Word:";
            PromptDialog.Choice(context, async (IDialogContext choiceContext, IAwaitable<string> choiceResult) =>
            {
                var selection = await choiceResult;
                OfficeOperation op = (OfficeOperation)Enum.Parse(typeof(OfficeOperation), selection);

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