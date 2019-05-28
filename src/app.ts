import * as express from "express";
import * as favicon from "serve-favicon";
import * as bodyParser from "body-parser";
import * as path from "path";
import * as logger from "winston";
import * as winston from "winston";
import * as msteams from "botbuilder-teams";
import config = require("config");
// initLogger();
winston.verbose("hello world");
// initLogger();
winston.verbose("hello world");
console.log(config["app"].baseUri);
let app = express();
app.set("port", process.env.PORT || 3333);
app.use(express.static(path.join(__dirname, "../../public")));
app.use(express.static(path.join(__dirname, "./views")));
app.use(favicon(path.join(__dirname, "../../public/images", "favicon.ico")));
app.use(bodyParser.json());
// Create bot
let connector = new msteams.TeamsChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
});
// NOTE: This endpoint cannot be changed and must be api/messages
app.post("/api/messages", connector.listen());
// Adding tabs to our app. This will setup routes to various views
let tabs = require("./tabs");
tabs.setup(app);
// Configure ping route
app.get("/ping", (req, res) => {
    res.status(200).send("OK");
});

// Start our nodejs app
app.listen(app.get("port"), function (): void {
    console.log(process.env.MicrosoftAppId);
    console.log(process.env.MicrosoftAppPassword);
    console.log("hello");
    console.log("Express server listening on port " + app.get("port"));
    console.log("Bot messaging endpoint: " + config["app"].baseUri + "/api/messages");
    logger.verbose("Express server listening on port " + app.get("port"));
    logger.verbose("Bot messaging endpoint: " + config["app"].baseUri + "/api/messages");
});

function initLogger(): void {
    logger.addColors({
        error: "red",
        warn: "yellow",
        info: "green",
        verbose: "cyan",
        debug: "blue",
        silly: "magenta",
    });

    logger.remove(logger.transports.Console);
    logger.add(logger.transports.Console, {
        timestamp: () => {
            return new Date().toLocaleTimeString();
        },
        colorize: process.env.MONOCHROME_CONSOLE ? false : true,
        prettyPrint: true,
        level: "debug",
    });
}
