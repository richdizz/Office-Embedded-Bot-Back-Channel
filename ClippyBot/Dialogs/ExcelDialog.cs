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
    public class ExcelDialog : IDialog<IMessageActivity>
    {
        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var msg = await result;
            
            string[] options = new string[] { "Chart", "Range" };
            string prompt = $"I see you are running the Clippy Bot in Excel. Select something and I'll insert it into a *NEW* worksheet:";
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