import * as express from "express";
import * as favicon from "serve-favicon";
import * as bodyParser from "body-parser";
import * as path from "path";
import * as logger from "winston";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as config from "config";

let app = express();
app.set("port", process.env.PORT || 3333);
app.use(express.static(path.join(__dirname, "../../public")));
app.use(express.static(path.join(__dirname, "./views")));
app.use(favicon(path.join(__dirname, "../../public/images", "favicon.ico")));
app.use(bodyParser.json());
// Create bot
let connector = new msteams.TeamsChatConnector({
  appId: config["app"].appId,
  appPassword: config["app"].appPassword,
});
// let bot = new builder.UniversalBot(connector);
let inMemoryStorage = new builder.MemoryBotStorage();
// NOTE: This endpoint cannot be changed and must be api/messages
app.post("/api/getAudioRecord", connector.listen());
app.post("/api/messages", connector.listen());
// Adding tabs to our app. This will setup routes to letious views
let tabs = require("./tabs");
tabs.setup(app);
// Configure ping route
app.get("/ping", (req, res) => {
  res.status(200).send("OK");
});
connector.resetAllowedTenants();

// Start our nodejs app
app.listen(app.get("port"), function(): void {
  console.log("Express server listening on port " + app.get("port"));
  console.log("Bot messaging endpoint: " + config["app"].baseUri + "/api/messages");
  logger.verbose("Express server listening on port " + app.get("port"));
  logger.verbose("Bot messaging endpoint: " + config["app"].baseUri + "/api/messages");
});
