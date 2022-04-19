import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";

export class ShowMail extends SSOCommand {
  constructor() {
    super();
    this.matchPatterns = [/^\s*mail\s*/];
    this.operationWithSSOToken = this.showUserInfo;
  }

  async showUserInfo(context: TurnContext, ssoToken: string) {
    await context.sendActivity("Retrieving mail information from Microsoft Graph ...");

    // Call Microsoft Graph half of user
    const teamsfx = new TeamsFx().setSsoToken(ssoToken);
    const graphClient = createMicrosoftGraphClient(teamsfx, [
      "User.Read","Mail.Read"
    ]);
    // get userprofile
    //const me = await graphClient.api("/me/messages").get();
    // get mail messages
    const mail = await graphClient.api("/me/messages")
          .select('sender,subject')
	        .get();

    if (mail) {
      //const value = mail.value;
      await context.sendActivity(
        `Mail information: ${JSON.stringify(mail)}`
      );
      // parse mail JSON
      const value = mail.value;
      const id = value[0].id;
      const sender = value[0].sender;
      const emailAddress = sender.emailAddress;
      await context.sendActivity(
        `Your first mail (${value[0].subject}) was sent by ${emailAddress.name}` 
        //was sent by ${emailAddress.name} (${emailAddress.address}).`
      );
      //await context.sendActivity({ attachments: [card] });
    }
    else {
      await context.sendActivity(
        "Could not retrieve mail messages from Microsoft Graph."
      );
    }
  }
}
