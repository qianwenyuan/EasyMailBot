import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";

export class MailHelp extends SSOCommand {
  constructor() {
    super();
    this.matchPatterns = [/^\s*mail\s*/];
    this.operationWithSSOToken = this.showMailHelp;
  }

  async showMailHelp(context: TurnContext, ssoToken: string) {
    // help information about commands
    await context.sendActivity(
        `Here are mail commands:\n
        1.Enter \"listrule\" to get several list rule templates.
        2.Enter \"rule+X(Exp:rule 1)\" to set the Xth template as your mail list rule.
        3.Enter \"listmail\" to fetch the 5 most recent mail under the rule you set(default all).`  
    );
  }
}
