import * as microsoftTeams from "@microsoft/teams-js";
import * as constants from "./constants";
import { appRoot } from "./dialogs/CardTemplates";
import { taskModuleLink } from "./utils/DeepLinks";

declare var appId: any; // Injected at template render time

// Set the desired theme
function setTheme(theme: string): void {
    if (theme) {
        // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
        document.body.className = "theme-" + (theme === "default" ? "light" : theme);
    }
}

// Create the URL that Microsoft Teams will load in the tab. You can compose any URL even with query strings.
function createTabUrl(): string {
    let tabChoice = document.getElementById("tabChoice");
    let selectedTab = tabChoice[(tabChoice as HTMLSelectElement).selectedIndex].value;

    return window.location.protocol + "//" + window.location.host + "/" + selectedTab;
}

// Call the initialize API first
microsoftTeams.initialize();

// Check the initial theme user chose and respect it
microsoftTeams.getContext(function(context: microsoftTeams.Context): void {
    if (context && context.theme) {
        setTheme(context.theme);
    }
});

// Handle theme changes
microsoftTeams.registerOnThemeChangeHandler(function(theme: string): void {
    setTheme(theme);
});

// Save configuration changes
microsoftTeams.settings.registerOnSaveHandler(function(saveEvent: microsoftTeams.settings.SaveEvent): void {
    microsoftTeams.settings.setSettings({
        contentUrl: createTabUrl(), // Mandatory parameter
        entityId: createTabUrl(), // Mandatory parameter
    });
    saveEvent.notifySuccess();
});
export function submittask(card: any): void {
    microsoftTeams.tasks.submitTask(card);
}
// Logic to let the user configure what they want to see in the tab being loaded
document.addEventListener("DOMContentLoaded", function(): void {
    let tabChoice = document.getElementById("tabChoice");
    if (tabChoice) {
        tabChoice.onchange = function(): void {
            let selectedTab = this[(this as HTMLSelectElement).selectedIndex].value;
            microsoftTeams.settings.setValidityState(selectedTab === "first" || selectedTab === "second" || selectedTab === "taskmodule");
        };
    }

    // If we are on the Task Module page, initialize the buttons and deep links
    let taskModuleButtons = document.getElementsByClassName("taskModuleButton");
    if (taskModuleButtons.length > 0) {
        // Initialize deep links
        let taskInfo = {
            title: null,
            height: null,
            width: null,
            url: null,
            card: null,
            fallbackUrl: null,
            completionBotId: null,
        };
        let deepLink = document.getElementById("dlaudioRecordApps") as HTMLAnchorElement;
        deepLink.href = taskModuleLink(constants.defaultJson.appId, constants.TaskModuleStrings.AudioRecordTitle, constants.TaskModuleSizes.audiorecord.height, constants.TaskModuleSizes.audiorecord.width, `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`, null, `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`);

        for (let btn of taskModuleButtons) {
            btn.addEventListener("click",
                function (): void {
                    taskInfo.url = `${appRoot()}/${this.id.toLowerCase()}`;
                    let submitHandler = (err: string, result: any): void => { console.log(`Err: ${err}; Result:  + ${result}`); };
                    taskInfo.title = constants.TaskModuleStrings.AudioRecordTitle;
                    taskInfo.height = constants.TaskModuleSizes.audiorecord.height;
                    taskInfo.width = constants.TaskModuleSizes.audiorecord.width;
                    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
                    });
                }
    }
});
