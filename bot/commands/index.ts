import { BotCommand } from "../helpers/botCommand";
import { LearnCommand } from "./learn";
import { ShowUserProfile } from "./showUserProfile";
import { WelcomeCommand } from "./welcome";
import { ShowMail } from "./showMail";
import { MailHelp } from "./mailHelp";
import { ListRules } from "./listRules";
import { SetRules } from "./setRules";

export const commands: BotCommand[] = [
  new LearnCommand(),
  new ShowUserProfile(),
  new WelcomeCommand(),
  new ShowMail(),
  new MailHelp(),
  new ListRules(),
  new SetRules()
]