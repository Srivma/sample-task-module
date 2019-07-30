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
        contentUrl: constants.TaskModuleInfo.contentUrl, // Mandatory parameter
        entityId: constants.TaskModuleInfo.taskName, // Mandatory parameter
    });
    saveEvent.notifySuccess();
});

// Logic to let the user configure what they want to see in the tab being loaded
document.addEventListener("DOMContentLoaded", function(): void {
    // If we are on the Task Module page, initialize the buttons and deep links
    let taskModuleButtons = document.getElementsByClassName("taskModuleButton");
    if (taskModuleButtons.length > 0) {
        // Initialize deep links
        let taskInfo = {
            title: null,
            height: null,
            width: null,
            url: null,
        };
        let deepLink = document.getElementById("dlaudioRecordApps") as HTMLAnchorElement;
        deepLink.href = taskModuleLink(appId, constants.TaskModuleStrings.AudioRecordTitle, constants.TaskModuleSizes.audiorecord.height, constants.TaskModuleSizes.audiorecord.width, `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`, null, `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`);
        let submitHandler = (err: string, result: any): void => { console.log(`Err: ${err}; Result:  + ${result}`); };

        for (let btn of taskModuleButtons) {
            btn.addEventListener("click",
                function (): void {
                    taskInfo.url = `${appRoot()}/${this.id.toLowerCase()}`;
                    taskInfo.title = constants.TaskModuleStrings.AudioRecordTitle;
                    taskInfo.height = constants.TaskModuleSizes.audiorecord.height;
                    taskInfo.width = constants.TaskModuleSizes.audiorecord.width;
                    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
                    });
                }
    }
});
