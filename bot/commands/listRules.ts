import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";
import { commonVar } from "./common";

export class ListRules extends SSOCommand {
  constructor() {
    super();
    this.matchPatterns = [/^\s*listrule\s*/];
    this.operationWithSSOToken = this.showMailHelp;
  }

  async showMailHelp(context: TurnContext, ssoToken: string) {
    // help information about commands
    await context.sendActivity(
        `Here are mail list rules:
        (Current rule type = ${commonVar.getRuleType()})\n
        rule 0: No rule. Get the most recent 5 mails.
        rule 1: Get all the mails whose importance=high.
        rule 2: Get all the mails sent by your teammates.
        `
    );
  }
}
