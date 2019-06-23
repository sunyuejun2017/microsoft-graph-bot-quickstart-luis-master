# Microsoft Graph Bot Quickstart w/ LUIS
This repository holds a quickstart template for building a bot that connects to the Microsoft Graph and leverages [LUIS](https://www.luis.ai/) for intents. A more simple quickstart for building bots with the Microsoft Graph **without LUIS** can be found [HERE](https://github.com/richdizz/microsoft-graph-bot-quickstart). Please follow the steps below to use this project.

**This is not designed for production. It exists primarily to get you started quickly with learning and prototyping bots that connect to the Microsoft Graph**

Please keep the above in mind before posting issues and PRs on this repo.

## Working with the Microsoft Graph in a Bot
This project uses the [BotAuth](https://github.com/richdizz/BotAuth) package(s) for handling authentication against Azure AD. BotAuth uses a magic number to secure the authentication flow, even when the bot is part of a group conversation. All of the logic for this Quickstart solution can be found in the [Dialogs/RootDialog.cs](https://github.com/richdizz/microsoft-graph-bot-quickstart-luis/blob/master/MsftGraphBotQuickStartLUIS/Dialogs/RootDialog.cs) file. Ultimately, you forward dialog to the AuthDialog, which handles the authentication flow and returns an access token for calling the Microsoft Graph.

## Steps for using the Quickstart
1. You will start by cloning the Quickstart template into a new project folder:

```
git clone https://github.com/richdizz/microsoft-graph-bot-quickstart-luis.git MyProjectName
```

2. Next, discard the git references:

```
rm -rf .git  # OS/X (bash)
rd .git /S/Q # windows
```

3. Open the project in Visual Studio and allow the NuGet packages to restore.

4. Create a Microsoft Cognitive Services key for LUIS from the [Azure Portal](https://portal.azure.com). For more information on how to do this, see [Creating Subscription Keys Via Azure](https://docs.microsoft.com/en-us/azure/cognitive-services/LUIS/azureibizasubscription).

5. Open the **Dialogs/RootDialog.cs** file and update the LuisModel settings with your subscription key.

    <pre>
    namespace MsftGraphBotQuickStart.Dialogs
    {
        [LuisModel("c138dfee-d228-416f-953e-5307c2b236ac", "COGNITIVE_SERVICES_SUBSCRIPTION_KEY_HERE")]
        [Serializable]
    </pre>

6. Build and run the project.

7. Launch the Bot Framework Emulator (available [HERE](https://docs.botframework.com/en-us/tools/bot-framework-emulator/)).

8. Enter the messaging endpoint for the bot project (likely http://localhost:3980/api/messages but could be a different port) and the click the "Connect" button (leave the Microsoft App ID and Microsoft App Password blank as these are for published bots).

9. Type anything to the bot and follow the prompts.

![Animated gif of Quickstart project run in the Bot Emulator](https://github.com/richdizz/microsoft-graph-bot-quickstart-luis/blob/master/Images/MsftGraphBotQuickStartLUIS.gif?raw=true)

## Going further
The first step in going further would be to register your own appliation at [https://apps.dev.office.com](https://apps.dev.office.com). This will allow you to play with different scopes and Microsoft Graph endpoints. You should also create your own LUIS model at [LUIS](https://www.luis.ai/). The LUIS model referenced in the code has been exported and added to the solution for your reference - [MsftGraphBotQuickStartLUIS.json](https://github.com/richdizz/microsoft-graph-bot-quickstart-luis/blob/master/MsftGraphBotQuickStartLUIS/MsftGraphBotQuickStartLUIS.json). It is also recommended you experiment with the bot in real Bot Framework channels (vs the emulator). The Bot Framework supports a number of channels, including Facebook Messenger, Microsoft Teams, Skype, and much more. You can register a bot at [https://bots.botframework.com](https://bots.botframework.com).

An easy addition to the solution would be to query thumbnails for the files that come back. You can read more about OneDrive's thumbnail API in the Microsoft Graph [HERE](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/item_list_thumbnails).