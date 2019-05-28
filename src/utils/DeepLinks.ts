// Helper function to generate task module deep links
import * as constants from "../constants";
export function taskModuleLink(
    appId: string,
    // tslint:disable:no-inferrable-types
    title: string = "",
    height: string | number = "medium",
    width: string | number = "medium",
    url: string = null,
    card: any = null,
    fallbackUrl?: string,
    completionBotId?: string): string {
        if ((url === null) && (card === null)) {
            return("Error generating deep link: you must specify either a card or URL.");
        }
        else {
            let cardOrUrl = (card === null) ? `url=${url}` : `card=${JSON.stringify(card)}`;
            let fallBack = (fallbackUrl === undefined) ? "" : `&fallbackUrl=${fallbackUrl}`;
            let completionBot = (completionBotId === undefined) ? "" : `&completionBotId=${constants.defaultJson.appId}`;
            return(encodeURI(`https://teams.microsoft.com/l/task/${constants.defaultJson.appId}?${cardOrUrl}&height=${height}&width=${width}&title=${title}${fallBack}${completionBot}`));
        }
}
