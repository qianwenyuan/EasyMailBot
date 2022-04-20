import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";

export class SetRules extends SSOCommand {
  constructor() {
    super();
    this.matchPatterns = [/^\s*rule\s*[0-9]\s*help\s*/];
    this.operationWithSSOToken = this.showMailHelp;
  }

  async showMailHelp(context: TurnContext, ssoToken: string) {
    // help information about commands
    await context.sendActivity(
        "Set your rules:\n"
    );
  }
}
