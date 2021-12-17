import * as Express from "express";
import * as http from "http";
import * as path from "path";
import * as morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import * as debug from "debug";
import * as compression from "compression";



// Initialize debug logging module
const log = debug("msteams");

log(`Initializing Microsoft Teams Express hosted App...`);

// Initialize dotenv, to use .env file settings if existing
// tslint:disable-next-line:no-var-requires
require("dotenv").config();



// The import of components has to be done AFTER the dotenv config
import * as allComponents from "./TeamsAppsComponents";

// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;

// Inject the raw request body onto the request object
// express.use(Express.json({
//     verify: (req, res, buf: Buffer, encoding: string): void => {
//         (req as any).rawBody = buf.toString();
//     }
// }));
express.use(Express.json({limit: "50mb"}));

express.use(Express.urlencoded({ extended: true }));

// Express configuration
express.set("views", path.join(__dirname, "/"));

// Add simple logging
express.use(morgan("tiny"));

// Add compression - uncomment to remove compression
express.use(compression());

// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsApiRouter(allComponents));

// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents
}));

// Set default web page
express.use("/", Express.static(path.join(__dirname, "web/"), {
    index: "index.html"
}));


express.use("/token",function(req, res) {
    const APP_ID = 'a013c4a4-0683-4125-8c05-4004c2c3cc6f';
const APP_SECERET = '5CVaG3r.a3y_OK2K9PRa39hy~lqc-HQC3a';
const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/08948d7c-43ee-4cae-9f2c-67e0464345d8/oauth2/v2.0/token';
const MS_GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

const axios = require('axios');
const qs = require('qs');

const postData = {
    client_id: APP_ID,
    scope: MS_GRAPH_SCOPE,
    client_secret: APP_SECERET,
    grant_type: 'client_credentials'
};
const header={
    "Access-Control-Allow-Origin":'*',
    'Access-Control-Allow-Headers': 'Origin, X-Requested-With, Content-Type, Accept, Authorization',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS'
}
axios.post(TOKEN_ENDPOINT, qs.stringify(postData),header)
        .then(response => {
            
            res.status(200).json({
                message: 'Access Tokens',
                accessToken: response.data.access_token
            })
          //return response.data.access_token
        })
        .catch(error => {
            //console.log(error);
        });
    }
)


// Set the port
express.set("port", port);

// Start the webserver
http.createServer(express).listen(port, () => {
    log(`Server running on ${port}`);
});
