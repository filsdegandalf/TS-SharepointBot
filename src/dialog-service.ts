import { QnAService } from "./qna-service";
import { AuthentificationService } from "./authentification-service";
import { QnAMakerDialog, QnAMakerRecognizer } from "botbuilder-cognitiveservices";
import { UniversalBot, Dialog, Prompts, Session, ChatConnector } from "botbuilder";

require("dotenv").config();

export class DialogService {
    connector: ChatConnector = new ChatConnector({
        appId: process.env.MICROSOFT_APP_ID,
        appPassword: process.env.MICROSOFT_APP_PASSWORD
    });
    bot: UniversalBot = new UniversalBot(this.connector);

    constructor() {
        this.basicDialog();
        this.authSuccessDialog();
        this.qnaDialog();
        console.log(this.bot);
    }


    /**
     * Base Dialog to start the chat !
     */
    basicDialog(): any {
        this.bot.dialog("/", [AuthentificationService.getAuthorization, (session) => {
            let keywords: string = session.message.text;

            if (keywords) {
                if (keywords === "reset") {
                    session.privateConversationData = {};
                    session.beginDialog("/");
                } else {
                    session.beginDialog("/qna");
                }
            }
        }
        ]);
    }

    /**
     * Call back of a successfull connection to SharePoint
     */
    authSuccessDialog(): any {
        this.bot.dialog("/oauth-success", (session: Session, response: any) => {
            if (encodeURIComponent(JSON.stringify(session.message.address)) !== encodeURIComponent(response.state)) {
                session.send("Fermeture de la session...");
                session.endDialog();
            } else {
                if (!session.privateConversationData.accessToken) {
                    session.privateConversationData.accessToken = response.accessToken;
                    session.privateConversationData.expiresOn = response.expiresOn;
                    session.privateConversationData.refreshToken = response.refreshToken;
                }

                session.send("Bonjour %s. Que recherchez vous?", response.userName);

                session.beginDialog("/");
            }
        });
    }

    /**
     * Dialog to use the qna service
     */
    qnaDialog(): any {
        this.bot.dialog("/qna", QnAService.basicQnAMakerDialog(this.bot));
    }


}