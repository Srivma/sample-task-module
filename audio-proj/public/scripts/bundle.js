(function(){function r(e,n,t){function o(i,f){if(!n[i]){if(!e[i]){var c="function"==typeof require&&require;if(!f&&c)return c(i,!0);if(u)return u(i,!0);var a=new Error("Cannot find module '"+i+"'");throw a.code="MODULE_NOT_FOUND",a}var p=n[i]={exports:{}};e[i][0].call(p.exports,function(r){var n=e[i][1][r];return o(n||r)},p,p.exports,r,e,n,t)}return n[i].exports}for(var u="function"==typeof require&&require,i=0;i<t.length;i++)o(t[i]);return o}return r})()({1:[function(require,module,exports){
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const microsoftTeams = require("@microsoft/teams-js");
const constants = require("./constants");
const CardTemplates_1 = require("./dialogs/CardTemplates");
const DeepLinks_1 = require("./utils/DeepLinks");
// Set the desired theme
function setTheme(theme) {
    if (theme) {
        // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
        document.body.className = "theme-" + (theme === "default" ? "light" : theme);
    }
}
// Create the URL that Microsoft Teams will load in the tab. You can compose any URL even with query strings.
function createTabUrl() {
    let tabChoice = document.getElementById("tabChoice");
    let selectedTab = tabChoice[tabChoice.selectedIndex].value;
    return window.location.protocol + "//" + window.location.host + "/" + selectedTab;
}
// Call the initialize API first
microsoftTeams.initialize();
// Check the initial theme user chose and respect it
microsoftTeams.getContext(function (context) {
    if (context && context.theme) {
        setTheme(context.theme);
    }
});
// Handle theme changes
microsoftTeams.registerOnThemeChangeHandler(function (theme) {
    setTheme(theme);
});
// Save configuration changes
microsoftTeams.settings.registerOnSaveHandler(function (saveEvent) {
    // Let the Microsoft Teams platform know what you want to load based on
    // what the user configured on this page
    microsoftTeams.settings.setSettings({
        contentUrl: createTabUrl(),
        entityId: createTabUrl(),
    });
    // Tells Microsoft Teams platform that we are done saving our settings. Microsoft Teams waits
    // for the app to call this API before it dismisses the dialog. If the wait times out, you will
    // see an error indicating that the configuration settings could not be saved.
    saveEvent.notifySuccess();
});
// Logic to let the user configure what they want to see in the tab being loaded
document.addEventListener("DOMContentLoaded", function () {
    // This module runs on multiple pages, so we need to isolate page-specific logic.
    // If we are on the tab configuration page, wire up the save button initialization state
    let tabChoice = document.getElementById("tabChoice");
    if (tabChoice) {
        tabChoice.onchange = function () {
            let selectedTab = this[this.selectedIndex].value;
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
        let deepLink = document.getElementById("dlaudioRecordApps");
        deepLink.href = DeepLinks_1.taskModuleLink(constants.defaultJson.appId, constants.TaskModuleStrings.AudioRecordTitle, constants.TaskModuleSizes.audiorecord.height, constants.TaskModuleSizes.audiorecord.width, `${CardTemplates_1.appRoot()}/${constants.TaskModuleIds.AudioRecord}`, null, `${CardTemplates_1.appRoot()}/${constants.TaskModuleIds.AudioRecord}`);
        for (let btn of taskModuleButtons) {
            btn.addEventListener("click", function () {
                taskInfo.url = `${CardTemplates_1.appRoot()}/${this.id.toLowerCase()}`;
                let submitHandler = (err, result) => { console.log(`Err: ${err}; Result:  + ${result}`); };
                taskInfo.title = constants.TaskModuleStrings.AudioRecordTitle;
                taskInfo.height = constants.TaskModuleSizes.audiorecord.height;
                taskInfo.width = constants.TaskModuleSizes.audiorecord.width;
                microsoftTeams.tasks.startTask(taskInfo, submitHandler);
            });
        }
    }
});



},{"./constants":2,"./dialogs/CardTemplates":3,"./utils/DeepLinks":4,"@microsoft/teams-js":5}],2:[function(require,module,exports){
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// Dialog ids
// tslint:disable-next-line:variable-name
exports.defaultJson = {
    baseUri: "https://sampleaudiorecordtask.azurewebsites.net",
    appId: "96f6fe44-166e-4951-b3a3-bb65ada49020",
    appPassword: ".1N-=7*yfK5Q24oP3iKF*P0Sj6jhZYMh",
};
// Task Module Strings
// tslint:disable-next-line:variable-name
exports.TaskModuleStrings = {
    AudioRecordTitle: "Create a new  audio record",
    AudioRecordName: "Audio Record",
};
// Task Module Ids
// tslint:disable-next-line:variable-name
exports.TaskModuleIds = {
    AudioRecord: "audiorecord",
};
// Task Module Sizes
// tslint:disable-next-line:variable-name
exports.TaskModuleSizes = {
    audiorecord: {
        width: 510,
        height: 255,
    },
};



},{}],3:[function(require,module,exports){
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const constants = require("../constants");
// Function that works both in Node (where window === undefined) or the browser
function appRoot() {
    return constants.defaultJson.baseUri;
}
exports.appRoot = appRoot;
exports.fetchTemplates = {
    // still implementing to get popup.
    audiorecord: {
        task: {
            type: "continue",
            value: {
                title: constants.TaskModuleStrings.AudioRecordTitle,
                height: constants.TaskModuleSizes.audiorecord.height,
                width: constants.TaskModuleSizes.audiorecord.width,
                fallbackUrl: `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`,
                url: `${appRoot()}/${constants.TaskModuleIds.AudioRecord}`,
            },
        },
    },
};



},{"../constants":2}],4:[function(require,module,exports){
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// Helper function to generate task module deep links
function taskModuleLink(appId, 
// tslint:disable:no-inferrable-types
title = "", height = "medium", width = "medium", url = null, card = null, fallbackUrl, completionBotId) {
    if ((url === null) && (card === null)) {
        return ("Error generating deep link: you must specify either a card or URL.");
    }
    else {
        let cardOrUrl = (card === null) ? `url=${url}` : `card=${JSON.stringify(card)}`;
        let fallBack = (fallbackUrl === undefined) ? "" : `&fallbackUrl=${fallbackUrl}`;
        let completionBot = (completionBotId === undefined) ? "" : `&completionBotId=${appId}`;
        return (encodeURI(`https://teams.microsoft.com/l/task/${appId}?${cardOrUrl}&height=${height}&width=${width}&title=${title}${fallBack}${completionBot}`));
    }
}
exports.taskModuleLink = taskModuleLink;



},{}],5:[function(require,module,exports){
!function(t,e){"function"==typeof define&&define.amd?define([],e):"object"==typeof exports?module.exports=e():t.microsoftTeams=e()}(this,function(){String.prototype.startsWith||(String.prototype.startsWith=function(t,e){return this.substr(!e||e<0?0:+e,t.length)===t});var t;return function(t){"use strict";function e(t){for(var e="^",n=t.split("."),i=0;i<n.length;i++)e+=(i>0?"[.]":"")+n[i].replace("*","[^/^.]+");return e+="$"}function n(t){for(var n="",i=0;i<t.length;i++)n+=(0===i?"":"|")+e(t[i]);return new RegExp(n)}function i(){if(!q){q=!0,R=this._window||window;var t=function(t){return k(t)};V=R.parent!==R.self?R.parent:R.opener,V?R.addEventListener("message",t,!1):(G=!0,window.onNativeMessage=T);try{_="*";var e=U(V,"initialize",[x]);tt[e]=function(t,e){X=t,Y=e}}finally{_=null}this._uninitialize=function(){X&&(r(null),s(null),c(null)),X===z.settings&&it.registerOnSaveHandler(null),X===z.remove&&it.registerOnRemoveHandler(null),G||R.removeEventListener("message",t,!1),q=!1,V=null,_=null,K=[],j=null,J=null,Q=[],Z=0,tt={},X=null,Y=null,G=!1}}}function o(t){b();var e=U(V,"getContext");tt[e]=t}function r(t){b(),$=t}function a(t){$&&$(t),j&&U(j,"themeChange",[t])}function s(t){b(),et=t}function u(t){et&&et(t)}function c(t){b(),nt=t}function f(){nt&&nt()||l()}function l(){b();var t=U(V,"navigateBack",[]);tt[t]=function(t){if(!t)throw new Error("Back navigation is not supported in the current client or context.")}}function h(t){b(z.content,z.settings,z.remove,z.task);var e=U(V,"navigateCrossDomain",[t]);tt[e]=function(t){if(!t)throw new Error("Cross-origin navigation is only supported for URLs matching the pattern registered in the manifest.")}}function v(t,e){b();var n=U(V,"getTabInstances",[e]);tt[n]=t}function d(t,e){b();var n=U(V,"getUserJoinedTeams",[e]);tt[n]=t}function g(t,e){b();var n=U(V,"getMruTabInstances",[e]);tt[n]=t}function p(t){b(z.content),U(V,"shareDeepLink",[t.subEntityId,t.subEntityLabel,t.subEntityWebUrl])}function m(t){b(z.content);var e=[t.entityId,t.title,t.description,t.type,t.objectUrl,t.downloadUrl,t.webPreviewUrl,t.webEditUrl,t.baseUrl,t.editFile,t.subEntityId];U(V,"openFilePreview",e)}function y(t){b();var e=U(V,"uploadCustomApp",[t]);tt[e]=function(t,e){if(!t)throw new Error(e)}}function w(t){b();var e=U(V,"navigateToTab",[t]);tt[e]=function(t){if(!t)throw new Error("Invalid internalTabInstanceId and/or channelId were/was provided")}}function b(){for(var t=[],e=0;e<arguments.length;e++)t[e]=arguments[e];if(!q)throw new Error("The library has not yet been initialized");if(X&&t&&t.length>0){for(var n=!1,i=0;i<t.length;i++)if(t[i]===X){n=!0;break}if(!n)throw new Error("This call is not allowed in the '"+X+"' context")}}function k(t){if(t&&t.data&&"object"==typeof t.data){var e=t.source||t.originalEvent.source,n=t.origin||t.originalEvent.origin;e===R||n!==R.location.origin&&!D.test(n.toLowerCase())||(C(e,n),e===V?T(t):e===j&&I(t))}}function C(t,e){V&&t!==V?j&&t!==j||(j=t,J=e):(V=t,_=e),V&&V.closed&&(V=null,_=null),j&&j.closed&&(j=null,J=null),S(V),S(j)}function T(t){if("id"in t.data){var e=t.data,n=tt[e.id];n&&(n.apply(null,e.args),delete tt[e.id])}else if("func"in t.data){var e=t.data,i=W[e.func];i&&i.apply(this,e.args)}}function I(t){if("id"in t.data&&"func"in t.data){var e=t.data,n=W[e.func];if(n){var i=n.apply(this,e.args);i&&O(j,e.id,Array.isArray(i)?i:[i])}else{var o=U(V,e.func,e.args);tt[o]=function(){for(var t=[],n=0;n<arguments.length;n++)t[n]=arguments[n];j&&O(j,e.id,t)}}}}function M(t){return t===V?K:t===j?Q:[]}function E(t){return t===V?_:t===j?J:null}function S(t){for(var e=E(t),n=M(t);t&&e&&n.length>0;)t.postMessage(n.shift(),e)}function N(t,e){var n=R.setInterval(function(){0===M(t).length&&(clearInterval(n),e())},100)}function U(t,e,n){var i=B(e,n);if(G)R&&R.nativeInterface&&R.nativeInterface.framelessPostMessage(JSON.stringify(i));else{var o=E(t);t&&o?t.postMessage(i,o):M(t).push(i)}return i.id}function A(t,e){return b(),U(V,t,e)}function O(t,e,n){var i=L(e,n),o=E(t);t&&o&&t.postMessage(i,o)}function B(t,e){return{id:Z++,func:t,args:e||[]}}function L(t,e){return{id:t,args:e||[]}}function H(t){b();var e=U(V,"getChatMembers");tt[e]=t}var P,x="1.3.7",F=["https://teams.microsoft.com","https://teams.microsoft.us","https://int.teams.microsoft.com","https://devspaces.skype.com","https://ssauth.skype.com","http://dev.local","https://msft.spoppe.com","https://*.sharepoint.com","https://*.sharepoint-df.com","https://*.sharepointonline.com","https://outlook.office.com","https://outlook-sdf.office.com"],D=n(F),W={},z={settings:"settings",content:"content",authentication:"authentication",remove:"remove",task:"task"};!function(t){function e(t,e){b(),l=e,U(V,"setUpViews",[t])}function n(t){l&&l(t)||(b(),U(V,"viewConfigItemPress",[t]))}function i(t,e){b(),c=e,U(V,"setNavBarMenu",[t])}function o(t){c&&c(t)||(b(),U(V,"handleNavBarMenuItemPress",[t]))}function r(t,e){b(),f=e,U(V,"showActionMenu",[t])}function a(t){f&&f(t)||(b(),U(V,"handleActionMenuItemPress",[t]))}var s=function(){function t(){this.enabled=!0}return t}();t.MenuItem=s;var u;!function(t){t.dropDown="dropDown",t.popOver="popOver"}(u=t.MenuListType||(t.MenuListType={}));var c;W.navBarMenuItemPress=o;var f;W.actionMenuItemPress=a;var l;W.setModuleView=n,t.setUpViews=e,t.setNavBarMenu=i,t.showActionMenu=r}(P=t.menus||(t.menus={}));var R,V,_,j,J,X,Y,$,q=!1,G=!1,K=[],Q=[],Z=0,tt={};W.themeChange=a;var et;W.fullScreenChange=u;var nt;W.backButtonPress=f,t.initialize=i,t.getContext=o,t.registerOnThemeChangeHandler=r,t.registerFullScreenHandler=s,t.registerBackButtonHandler=c,t.navigateBack=l,t.navigateCrossDomain=h,t.getTabInstances=v,t.getUserJoinedTeams=d,t.getMruTabInstances=g,t.shareDeepLink=p,t.openFilePreview=m,t.uploadCustomApp=y,t.navigateToTab=w;var it;!function(t){function e(t){b(z.settings,z.remove),U(V,"settings.setValidityState",[t])}function n(t){b(z.settings,z.remove);var e=U(V,"settings.getSettings");tt[e]=t}function i(t){b(z.settings),U(V,"settings.setSettings",[t])}function o(t){b(z.settings),u=t}function r(t){b(z.remove),c=t}function a(t){var e=new f(t);u?u(e):e.notifySuccess()}function s(){var t=new l;c?c(t):t.notifySuccess()}var u,c;W["settings.save"]=a,W["settings.remove"]=s,t.setValidityState=e,t.getSettings=n,t.setSettings=i,t.registerOnSaveHandler=o,t.registerOnRemoveHandler=r;var f=function(){function t(t){this.notified=!1,this.result=t?t:{}}return t.prototype.notifySuccess=function(){this.ensureNotNotified(),U(V,"settings.save.success"),this.notified=!0},t.prototype.notifyFailure=function(t){this.ensureNotNotified(),U(V,"settings.save.failure",[t]),this.notified=!0},t.prototype.ensureNotNotified=function(){if(this.notified)throw new Error("The SaveEvent may only notify success or failure once.")},t}(),l=function(){function t(){this.notified=!1}return t.prototype.notifySuccess=function(){this.ensureNotNotified(),U(V,"settings.remove.success"),this.notified=!0},t.prototype.notifyFailure=function(t){this.ensureNotNotified(),U(V,"settings.remove.failure",[t]),this.notified=!0},t.prototype.ensureNotNotified=function(){if(this.notified)throw new Error("The removeEvent may only notify success or failure once.")},t}()}(it=t.settings||(t.settings={}));var ot;!function(t){function e(t){g=t}function n(t){var e=void 0!==t?t:g;if(b(z.content,z.settings,z.remove,z.task),"desktop"===Y){var n=document.createElement("a");n.href=e.url;var i=U(V,"authentication.authenticate",[n.href,e.width,e.height]);tt[i]=function(t,n){t?e.successCallback(n):e.failureCallback(n)}}else a(e)}function i(t){b();var e=U(V,"authentication.getAuthToken",[t.resources]);tt[e]=function(e,n){e?t.successCallback(n):t.failureCallback(n)}}function o(t){b();var e=U(V,"authentication.getUser");tt[e]=function(e,n){e?t.successCallback(n):t.failureCallback(n)}}function r(){s();try{j&&j.close()}finally{j=null,J=null}}function a(t){g=t,r();var e=g.width||600,n=g.height||400;e=Math.min(e,R.outerWidth-400),n=Math.min(n,R.outerHeight-200);var i=document.createElement("a");i.href=g.url;var o="undefined"!=typeof R.screenLeft?R.screenLeft:R.screenX,a="undefined"!=typeof R.screenTop?R.screenTop:R.screenY;o+=R.outerWidth/2-e/2,a+=R.outerHeight/2-n/2,j=R.open(i.href,"_blank","toolbar=no, location=yes, status=no, menubar=no, scrollbars=yes, top="+a+", left="+o+", width="+e+", height="+n),j?u():h("FailedToOpenWindow")}function s(){p&&(clearInterval(p),p=0),delete W.initialize,delete W.navigateCrossDomain}function u(){s(),p=R.setInterval(function(){if(!j||j.closed)h("CancelledByUser");else{var t=J;try{J="*",U(j,"ping")}finally{J=t}}},100),W.initialize=function(){return[z.authentication,Y]},W.navigateCrossDomain=function(t){return!1}}function c(t,e){v(e,"result",t),b(z.authentication),U(V,"authentication.authenticate.success",[t]),N(V,function(){return setTimeout(function(){return R.close()},200)})}function f(t,e){v(e,"reason",t),b(z.authentication),U(V,"authentication.authenticate.failure",[t]),N(V,function(){return setTimeout(function(){return R.close()},200)})}function l(t){try{g&&g.successCallback&&g.successCallback(t)}finally{g=null,r()}}function h(t){try{g&&g.failureCallback&&g.failureCallback(t)}finally{g=null,r()}}function v(t,e,n){if(t){var i=document.createElement("a");i.href=decodeURIComponent(t),i.host&&i.host!==window.location.host&&"outlook.office.com"===i.host&&i.search.indexOf("client_type=Win32_Outlook")>-1&&(e&&"result"===e&&(n&&(i.href=d(i.href,"result",n)),R.location.assign(d(i.href,"authSuccess",""))),e&&"reason"===e&&(n&&(i.href=d(i.href,"reason",n)),R.location.assign(d(i.href,"authFailure",""))))}}function d(t,e,n){var i=t.indexOf("#"),o=i===-1?"#":t.substr(i);return o=o+"&"+e+(""!==n?"="+n:""),t=i===-1?t:t.substr(0,i),t+o}var g,p;W["authentication.authenticate.success"]=l,W["authentication.authenticate.failure"]=h,t.registerAuthenticationHandlers=e,t.authenticate=n,t.getAuthToken=i,t.getUser=o,t.notifySuccess=c,t.notifyFailure=f}(ot=t.authentication||(t.authentication={})),t.sendCustomMessage=A;var rt;!function(t){function e(t,e){b(z.content);var n=U(V,"tasks.startTask",[t]);tt[n]=e}function n(t,e){b(z.content,z.task),U(V,"tasks.completeTask",[t,Array.isArray(e)?e:[e]])}t.startTask=e,t.submitTask=n}(rt=t.tasks||(t.tasks={})),t.getChatMembers=H}(t||(t={})),t});
},{}]},{},[1]);