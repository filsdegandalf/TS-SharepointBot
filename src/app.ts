import { DialogService } from "./dialog-service";
import { AuthentificationService } from "./authentification-service";
import { Server } from "restify";
import { UniversalBot } from "botbuilder";
import * as restify from "restify";
import * as builder from "botbuilder";
import { ChatConnector } from "botbuilder";

require("dotenv").config();
let bot: UniversalBot = new DialogService().bot;


let port: string = process.env.port || process.env.PORT || '3978';
const server: Server = restify.createServer();

// change done for restify 5.X+ (mapParams should be specified @ true)
server.use(restify.plugins.bodyParser({
    mapParams: true
}));
server.listen(port, () => console.log(`${server.name} listening to ${server.url}`));

// entry point of your bot
server.post("/api/messages", (bot.connector("*") as builder.ChatConnector).listen());

server.post("api/oauthcallback", (req: restify.Request, res: restify.Response, next) => {
    let authorizationCode: string = req.params.code;

    if (authorizationCode !== undefined) {
        AuthentificationService.acquireTokenWithAuthorizationCode(authorizationCode).then((response) => {
            let state: any = req.params.state;

            if (state) {
                let address: any = JSON.parse(state);
                response.state = state;
                bot.beginDialog(address, "/oauth-success", response);
            }
            let body: string = "<html><body>Authentication succeeded! You can now close this tab</body></html>";
            res.send(200, body, { "Content-Length": Buffer.byteLength(body).toString(), "Content-Type": "text/html" });
            res.end();

        }).catch((errorMessage) => {

            var body: string = "<html><body>" + errorMessage + "</body></html>";
            res.send(200, body, { "Content-Length": Buffer.byteLength(body).toString(), "Content-Type": "text/html" });
            res.end();
        });
    } else {
        var body: string = "<html><body>" + "unable to retrieve the authentication code" + "</body></html > ";
        res.send(200, body, { "Content-Length": Buffer.byteLength(body).toString(), "Content-Type": "text/html" });
        res.end();
    }
});
