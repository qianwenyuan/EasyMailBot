import { ResponseType } from "@microsoft/microsoft-graph-client";
import { CardFactory, TurnContext } from "botbuilder";
import {
  createMicrosoftGraphClient,
  TeamsFx,
} from "@microsoft/teamsfx";
import { SSOCommand } from "../helpers/botCommand";
import { commonVar } from "./common";

function getType(userInput: string) {
  var n=userInput.match(/[0-9]/);
  if (n) return n[0];
  return -1;
}
export class SetRules extends SSOCommand {
  ruleType;
  constructor() {
    super();
  this.matchPatterns = [/^\s*setrule\s[0-9]\s*/];
    this.operationWithSSOToken = this.showMailHelp;
  }
  //getInput(userInput) {this.userInput=userInput;}

  async showMailHelp(context: TurnContext, ssoToken: string) {
    await context.sendActivity("Setting rules...");

    this.ruleType = getType(commonVar.userInput);
    if (this.ruleType>=0&&this.ruleType<=2)
      commonVar.setRuleType(this.ruleType);
    // help information about commands
    if (this.ruleType>=0&&this.ruleType<=2) {
      await context.sendActivity(
        `Set your rules as rule ${this.ruleType}`
      );
    }
    else {
      await context.sendActivity(
        `No such rule type! Please contact the administrator to add a rule.`
      )
    }
  }
}
