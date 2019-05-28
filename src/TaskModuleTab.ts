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
    // Let the Microsoft Teams platform know what you want to load based on
    // what the user configured on this page
    microsoftTeams.settings.setSettings({
        contentUrl: createTabUrl(), // Mandatory parameter
        entityId: createTabUrl(), // Mandatory parameter
    });

    // Tells Microsoft Teams platform that we are done saving our settings. Microsoft Teams waits
    // for the app to call this API before it dismisses the dialog. If the wait times out, you will
    // see an error indicating that the configuration settings could not be saved.
    saveEvent.notifySuccess();
});

// Logic to let the user configure what they want to see in the tab being loaded
document.addEventListener("DOMContentLoaded", function(): void {
    // This module runs on multiple pages, so we need to isolate page-specific logic.

    // If we are on the tab configuration page, wire up the save button initialization state
    let tabChoice = document.getElementById("tabChoice");
    if (tabChoice) {
        tabChoice.onchange = function(): void {
            let selectedTab = this[(this as HTMLSelectElement).selectedIndex].value;

            // This API tells Microsoft Teams to enable the 'Save' button. Since Microsoft Teams always assumes
            // an initial invalid state, without this call the 'Save' button will never be enabled.
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
        deepLink.href = taskModuleLink(appId, constants.TaskModuleStrings.AudioRecordTitle, constants.TaskModuleSizes.audiorecord.height, constants.TaskModuleSizes.audiorecord.width, `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`, null, `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`);

        for (let btn of taskModuleButtons) {
            btn.addEventListener("click",
                function (): void {
                    taskInfo.url = `${appRoot()}/${this.id.toLowerCase()}`;
                    let submitHandler = (err: string, result: any): void => { console.log(`Err: ${err}; Result:  + ${result}`); };
                    switch (this.id.toLowerCase()) {
                        case constants.TaskModuleIds.AudioRecord:
                            taskInfo.title = constants.TaskModuleStrings.AudioRecordTitle;
                            taskInfo.height = constants.TaskModuleSizes.audiorecord.height;
                            taskInfo.width = constants.TaskModuleSizes.audiorecord.width;
                            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
                            break;
                        default:
                            console.log("Unexpected button ID: " + this.id.toLowerCase());
                            return;
                    }
                    console.log("URL: " + taskInfo.url);
                });
        }
    }
});
