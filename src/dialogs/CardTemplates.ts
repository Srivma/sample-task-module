import * as constants from "../constants";
// Function that works both in Node (where window === undefined) or the browser
export function appRoot(): string {
    return constants.defaultJson.baseUri;
    // return "https://02be80eb.ngrok.io";
}
export const fetchTemplates: any = {
    audiorecord: {
        task: {
            type: "continue",
            value: {
                title: constants.TaskModuleStrings.AudioRecordTitle,
                height: constants.TaskModuleSizes.audiorecord.height,
                width: constants.TaskModuleSizes.audiorecord.width,
                fallbackUrl: "${appRoot()}/${constants.TaskModuleIds.AudioRecord}",
                url: "${appRoot()}/${constants.TaskModuleIds.AudioRecord}",
            },
        },
    },
};
