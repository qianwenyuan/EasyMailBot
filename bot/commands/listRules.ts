import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";

export class ListRules extends SSOCommand {
  constructor() {
    super();
    this.matchPatterns = [/^\s*listrule\s*help\s*/];
    this.operationWithSSOToken = this.showMailHelp;
  }

  async showMailHelp(context: TurnContext, ssoToken: string) {
    // help information about commands
    await context.sendActivity(
        "Here are mail list rules:\n"+
        "rule 1: Get all the mails whose importance=high in past 48 hours.\n"+
        "rule 2: Get all the mails sent by your teammates in past 48 hours.\n"
    );
  }
}
