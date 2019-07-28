import * as express from "express";
import * as favicon from "serve-favicon";
import * as bodyParser from "body-parser";
import * as path from "path";
import * as logger from "winston";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as config from "config";
import { fetchTemplates } from "./dialogs/CardTemplates";
import * as azure from "azure-storage";
import * as stream from "stream";

let app = express();
app.set("port", process.env.PORT || 3334);
app.use(express.static(path.join(__dirname, "../../public")));
app.use(express.static(path.join(__dirname, "./views")));
app.use(favicon(path.join(__dirname, "../../public/images", "favicon.ico")));
app.use(bodyParser.json());

let connector = new msteams.TeamsChatConnector({
  appId: config.get("app.appId"),
  appPassword: config.get("app.appPassword"),
});
app.post("/api/messages", connector.listen());
// Adding tabs to our app. This will setup routes to views
// get all todos

let tabs = require("./tabs");
tabs.setup(app);
connector.resetAllowedTenants();
/* tslint:disable:typedef */
/* tslint:disable:typedef-whitespace */
let accountName = "audiorecordstorage";
let accountKey = config["app"].accountKey;
let blobService = azure.createBlobService(accountName, accountKey);
blobService.createContainerIfNotExists("taskcontainer", {
  publicAccessLevel: "blob",
}, function(error, result, response){
  if (!error) {
    console.log(result);
    console.log(response);
  }
});

app.use(function(req, res, next) {
  let data: any;
  let bufData : Buffer;
  let i = 0;
  let arr = [];
  req.on("data", function(chunk) {
    data = Buffer.from(chunk);
    arr.push(data);
    bufData = Buffer.concat(arr);
  });
  req.on("end", function() {
    let bufToStream = bufferToStream(bufData);
    req.body = bufToStream;
    next();
  });
});

app.use(bodyParser.raw({ type: "audio/wav", limit: "50mb" }));

function bufferToStream(buffer: any) {
  let Duplex = stream.Duplex;
  let streamData = new Duplex();
  streamData.push(buffer);
  streamData.push(null);
  // streamData.pipe(process.stdout);
  return streamData;
}

app.post("/audiodata", (req, res) => {
  let date = new Date().getTime();
  let blobName = "testblob" + date;
  let downloadLink;
  let file: number;
  file = Number(req.headers["content-length"]);
  blobService.createBlockBlobFromStream("audiofilecontainer",
  blobName,
  req.body,
  file,
  (error, result) => {
    if (error) {
      console.log("Couldn't upload file %s", error);
    } else {
      downloadLink = blobService.getUrl("audiofilecontainer", blobName);
      res.status(200).send(downloadLink);
    }
  });
});

// Start our nodejs app
app.listen(app.get("port"), function(): void {
  console.log("Express server listening on port " + app.get("port"));
  console.log("Bot messaging endpoint: " + config.get("app.baseUri") + "/api/messages");
  logger.verbose("Express server listening on port " + app.get("port"));
  logger.verbose("Bot messaging endpoint: " + config.get("app.baseUri") + "/api/messages");
});

/* tslint:disable:typedef */
/* tslint:disable:typedef-whitespace */
async function composeInvoke(event: builder.IEvent, cb: (err: Error, body: any, status?: number) => void) {
  let attachments = [];
  if ( event["name"] === "composeExtension/fetchTask" ) {
  cb(null, fetchTemplates["audiorecord"], 200);
  }
  if ( event["name"] === "composeExtension/submitAction" ) {
  let invokeValue = event["value"].data;
  let actionUrl = invokeValue.url;
  fetchTemplates["audiofile"].task["value"].url = "https://e88c5c14.ngrok.io/playaudio?audioFile=" + actionUrl;
  attachments.push(new builder.HeroCard()
  .title("Audio Card")
  .text("This card can play an audio file.")
  .buttons([builder.CardAction.invoke(null, "type", "task/fetch", "PlayAudio")])
  .toAttachment());
  let response = msteams.ComposeExtensionResponse.result("list").attachments(attachments).toResponse();
  cb(null, response, 200);
  }
  if ( event["name"] === "task/fetch" ) {
    // fetchTemplates["audiofile"].task["value"].url = "/playaudio?audiourl=" + actionUrl
    cb(null, fetchTemplates["audiofile"], 200);
  }
}

connector.onInvoke(composeInvoke);
