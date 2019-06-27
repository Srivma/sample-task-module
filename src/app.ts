import * as express from "express";
import * as favicon from "serve-favicon";
import * as bodyParser from "body-parser";
import * as path from "path";
import * as logger from "winston";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as config from "config";
import * as faker from "faker";
import { fetchTemplates } from "./dialogs/CardTemplates";

let app = express();
app.set("port", process.env.PORT || 3334);
app.use(express.static(path.join(__dirname, "../../public")));
app.use(express.static(path.join(__dirname, "./views")));
app.use(favicon(path.join(__dirname, "../../public/images", "favicon.ico")));
app.use(bodyParser.json());
// Create bot
let tests: any;
let connector = new msteams.TeamsChatConnector({
  appId: "96f6fe44-166e-4951-b3a3-bb65ada49020",
  appPassword: "drvIXmVag/y1zY**548WydGguTaPrG:m",
});
let inMemoryStorage = new builder.MemoryBotStorage();
let bot = new builder.UniversalBot(connector, {
  storage: new builder.MemoryBotStorage()});
app.post("/api/messages", connector.listen());
// Adding tabs to our app. This will setup routes to letious views
let tabs = require("./tabs");
tabs.setup(app);
connector.resetAllowedTenants();
// Start our nodejs app
app.listen(app.get("port"), function(): void {
  console.log("Express server listening on port " + app.get("port"));
  console.log("Bot messaging endpoint: " + config["app"].appId + "/api/messages");
  logger.verbose("Express server listening on port " + config["app"].appPassword);
  logger.verbose("Bot messaging endpoint: " + config["app"].appId + "/api/messages");
});
// we can use this below function in future
/*let composeExtensionHandler = function (event: builder.IEvent, query: msteams.ComposeExtensionQuery, callback: any): void {
  let title = query.parameters && query.parameters[0].name === "audioTitle"
  ? query.parameters[0].value
  : faker.lorem.sentence();
  let attachments = [];
  let randomImageUrl = "https://loremflickr.com/200/200";
  // Faker's random images uses lorempixel.com, which has been down a lot
  // Generate 5 results to send with fake text and fake images
  for (let i = 0; i < 5; i++) {
    attachments.push(
      new builder.HeroCard()
      .title(title)
      .text(faker.lorem.paragraph())
      .images([new builder.CardImage().url(`${randomImageUrl}?random=${i}`)])
      .toAttachment());
    }
  let response = msteams.ComposeExtensionResponse.result("list").attachments(attachments).toResponse();
  callback(null, response, 200);
};
function sendMessage (address: any, message: any) {
  bot.loadSession(address, function (e, session: builder.Session) {
      // session.send(message);
      session.send("hello");
  });
}
connector.onQuery("getAudioRecord", composeExtensionHandler);*/
/* tslint:disable:typedef */
/* tslint:disable:typedef-whitespace */
async function composeInvoke(event: builder.IEvent, cb: (err: Error, body: any, status?: number) => void) {
  let attachments = [];
  if ( event["name"] === "composeExtension/fetchTask" ) {
  cb(null, fetchTemplates["audiorecord"], 200);
  }
  if ( event["name"] === "composeExtension/submitAction" ) {
  attachments.push(new builder.HeroCard()
  .title("test")
  .text("The Empire Strikes Back (also known as Star Wars: Episode V – The Empire Strikes Back) is a 1980 American epic space opera film directed by Irvin Kershner. Leigh Brackett and Lawrence Kasdan wrote the screenplay, with George Lucas writing the film\'s story and serving as executive producer. The second installment in the original Star Wars trilogy, it was produced by Gary Kurtz for Lucasfilm Ltd. and stars Mark Hamill, Harrison Ford, Carrie Fisher, Billy Dee Williams, Anthony Daniels, David Prowse, Kenny Baker, Peter Mayhew and Frank Oz.")
  .buttons([builder.CardAction.openUrl(null, "https://19febf26.ngrok.io/test.wav", "Audio")])
  .toAttachment());
  let response = msteams.ComposeExtensionResponse.result("list").attachments(attachments).toResponse();
  cb(null, response, 200);
}}
connector.onInvoke(composeInvoke);
