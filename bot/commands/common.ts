export class CommonVar {
    userInput;
    ruleType;
    constructor() {
        this.userInput="";
        this.ruleType=0;
    }

    getUserInput() {return this.userInput;}
    setUserInput(userInput) {this.userInput=userInput;}

    getRuleType() {return this.ruleType;}
    setRuleType(ruleType) {this.ruleType=ruleType;}
}

export const commonVar = new CommonVar();