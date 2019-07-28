import { Request, Response } from "express";
import * as bodyParser from "body-parser";
import { appRoot } from "./dialogs/CardTemplates";

module.exports.setup = function(app: any): void {
    let path = require("path");
    // Configure the view engine, views folder and the statics path
    app.set("view engine", "pug");
    app.set("views", path.join(__dirname, "views"));

    app.use(bodyParser.json());
    app.use(bodyParser.urlencoded({
        extended: true,
    }));

    app.get("/taskmodule", function(req: Request, res: Response): void {
        // Render the template, passing the appId so it's included in the rendered HTML
        res.render("taskmodule", { appId: `${appRoot()}`});
    });

    app.get("/audiorecord", function(req: Request, res: Response): void {
        // Render the template, passing the appId so it's included in the rendered HTML
        res.render("audiorecord", { appId: `${appRoot()}` });
    });

    app.get("/playaudio", function(req: Request, res: Response): void {
        // Render the template, passing the appId so it's included in the rendered HTML
        res.render("playaudio", { appId: `${appRoot()}` });
    });
};
