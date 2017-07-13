
Registering with the Microsoft Bot Framework

Head over to https://dev.botframework.com/, click Register a Bot.
Fill in a name and a “bot handle” – this will be your BotId in your web.config – and paste in your Azure endpoint + “/api/messages/” for the “Messaging endpoint” section.

Then click “Create Microsoft App ID and Password” to open up a new tab and log in with your Microsoft (Live/whatever it’s called now) details to see your already configured apps.

Click “Add an app”, give it a name, and on the next screen you’ll see your Microsoft App ID; paste this back in the bot registration page AND in your web.config as the MicrosoftAppId.

After registration:
 Open web.config file and update these fields.
  <appSettings>
    <!-- update these with your BotId, Microsoft App Id and your Microsoft App Password-->
    <add key="BotId" value="KlippyBot" />
    <add key="MicrosoftAppId" value="XXXXXX" />
    <add key="MicrosoftAppPassword" value="YYYYYY" />
  </appSettings>
