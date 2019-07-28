import * as constants from "../constants";
import * as config from "config";
// Function that works both in Node (where window === undefined) or the browser
export function appRoot(): string {
    if (typeof window === "undefined") {
        return config.get("app.baseUri");
    }
    else {
        return window.location.protocol + "//" + window.location.host;
    }
}

export const fetchTemplates: any = {
    audiorecord: {
        task: {
            type: "continue",
            value: {
                title: constants.TaskModuleStrings.AudioRecordTitle,
                height: constants.TaskModuleSizes.audiorecord.height,
                width: constants.TaskModuleSizes.audiorecord.width,
                fallBackUrl: `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`,
                url: `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`,
            },
        },
    },
    audiofile: {
        task: {
            type: "continue",
            value: {
                title: constants.TaskModuleStrings.PlayAudioTitle,
                height: constants.TaskModuleSizes.playaudio.height,
                width: constants.TaskModuleSizes.playaudio.width,
                fallBackUrl: `${appRoot()}/${constants.TaskModuleIds.PlayAudio}`,
                url: `${appRoot()}/${constants.TaskModuleIds.PlayAudio}`,
            },
        },
    },
};
