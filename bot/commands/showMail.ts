import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";
import { commonVar } from "./common";

export class ShowMail extends SSOCommand {
  constructor() {
    super();
  this.matchPatterns = [/^\s*listmail\s*/];
    this.operationWithSSOToken = this.showUserInfo;
  }

  async showUserInfo(context: TurnContext, ssoToken: string) {
    await context.sendActivity("Retrieving mail information from Outlook ...");

    // Call Microsoft Graph half of user
    const teamsfx = new TeamsFx().setSsoToken(ssoToken);
    const graphClient = createMicrosoftGraphClient(teamsfx, [
      "User.Read","Mail.Read"
    ]);
    // get userprofile
    //const me = await graphClient.api("/me/messages").get();
    // get mail messages
    // dont fetch all mails;
    const mail = await graphClient.api("/me/messages?$top=10")
          .select('sender,receivedDateTime,importance,isRead,bodyPreview,subject')
	        .get();

    if (mail) {
      // parse mail JSON
      const value = mail.value;
      if (commonVar.getRuleType()==1) { // rule 1: importance
        await context.sendActivity(`Under list rule 1:`)
        for (var mailnum=0, validnum=0; mailnum<value.length && validnum<5; mailnum++) {
          const id = value[mailnum].id;
          const sender = value[mailnum].sender;
          const emailAddress = sender.emailAddress;
          const importance = value[mailnum].importance;
          const bodyPreview = value[mailnum].bodyPreview;
          const receivedDateTime = value[mailnum].receivedDateTime;

          if (importance=="high") {
            validnum++;
            await context.sendActivity(
              `mail ${validnum}: \'${value[mailnum].subject}\' was sent by ${emailAddress.name} at ${receivedDateTime}\n
              Here is the preview: ${bodyPreview}` 
            //was sent by ${emailAddress.name} (${emailAddress.address}).`
            );
          }
        }
      }
      else if (commonVar.getRuleType()==2) { // rule 2: is Read & reverse time order
        await context.sendActivity(`Under list rule 2:`)
        for (var mailnum=value.length-1, validnum=0; mailnum>=0 && validnum<5; mailnum--) {
          const id = value[mailnum].id;
          const sender = value[mailnum].sender;
          const emailAddress = sender.emailAddress;
          const isRead = value[mailnum].isRead;
          const bodyPreview = value[mailnum].bodyPreview;
          const receivedDateTime = value[mailnum].receivedDateTime;

          if (!isRead) {
            validnum++;
            await context.sendActivity(
              `mail ${validnum}: \'${value[mailnum].subject}\' was sent by ${emailAddress.name} at ${receivedDateTime}\n
              Here is the preview: ${bodyPreview}` 
              //was sent by ${emailAddress.name} (${emailAddress.address}).`
            );
          }
        }
      }
      else if (commonVar.getRuleType()==0) { // rule 0: No rule
        await context.sendActivity(`Under list rule 0:`)
        for (var mailnum=0; mailnum<value.length && mailnum<5; mailnum++) {
          const id = value[mailnum].id;
          const sender = value[mailnum].sender;
          const emailAddress = sender.emailAddress;
          const bodyPreview = value[mailnum].bodyPreview;
          const receivedDateTime = value[mailnum].receivedDateTime;

          await context.sendActivity(
            `mail ${mailnum+1}: \'${value[mailnum].subject}\' was sent by ${emailAddress.name} at ${receivedDateTime}\n
            Here is the preview: ${bodyPreview}` 
            //was sent by ${emailAddress.name} (${emailAddress.address}).`
          );
        } 
      }
    }
    else {
      await context.sendActivity(
        "Could not retrieve mail messages from Microsoft Graph."
      );
    }
  }
}
