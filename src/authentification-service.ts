import { CardAction, ICardAction, SigninCard, ThumbnailCard, Message, Session } from "botbuilder";
import { AuthenticationContext } from "adal-node";

export class AuthentificationService {
    private static adalConfig: any;

    constructor() {
        AuthentificationService.adalConfig = AuthentificationService.getAdalConfig();
    }

    /**
     * Retrieve the Token with the authorization code retrieved from request parameters
     */
    public static acquireTokenWithAuthorizationCode(authorizationCode: string): Promise<any> {
        AuthentificationService.checkAdal();
        let authenticationContext: AuthenticationContext = new AuthenticationContext(this.adalConfig.authorityUrl);

        let p: Promise<any> = new Promise((resolve, reject) => {

            authenticationContext.acquireTokenWithAuthorizationCode(
                authorizationCode,
                this.adalConfig.redirectUri,
                this.adalConfig.resource,
                this.adalConfig.clientId,
                this.adalConfig.clientSecret,
                (err, response: any) => {

                    if (err) {
                        reject("error: " + err.message + "\n");

                    } else {
                        resolve({
                            userName: (response.givenName + " " + response.familyName),
                            accessToken: response.accessToken,
                            expiresOn: response.expiresOn,
                            refreshToken: response.refreshToken,
                        });
                    }
                });
        });

        return p;
    }

    /**
     * Retrieve the Token with the refresh token retrieved from the bot session
     */
    public static acquireTokenWithRefreshToken(refreshToken: string): Promise<any> {
        AuthentificationService.checkAdal();
        let authenticationContext: AuthenticationContext = new AuthenticationContext(this.adalConfig.authorityUrl);

        let p: Promise<any> = new Promise((resolve, reject) => {
            authenticationContext.acquireTokenWithRefreshToken(
                refreshToken,
                this.adalConfig.clientId,
                this.adalConfig.clientSecret,
                this.adalConfig.resource,
                (err, response: any) => {
                    if (err) {
                        reject("error : " + err.message + "\n");
                    } else {
                        resolve({
                            userName: (response.givenName + " " + response.familyName),
                            accessToken: response.accessToken,
                            expiresOn: response.expiresOn,
                            refreshToken: response.refreshToken,
                        });
                    }
                }
            );
        });

        return p;
    }

    /**
     * Get authorization to connect to Sharepoint
     */
    public static getAuthorization(session: Session, args: any, next: Function): any {
        AuthentificationService.checkAdal();

        // if user is not authentified
        if (!session.privateConversationData.accessToken) {
            let stateToken: string = encodeURIComponent(JSON.stringify(session.message.address));
            let authorizationUrl: string = AuthentificationService.adalConfig.templateAuthzUrl.replace("<state>", stateToken);

            let actionLabel: string = "Vous devez vous connecter avant d'utiliser ce bot";
            let buttonLabel: string = "Se connecter";
            let signinCard: any = null;

            // the Sign-In card is not supported by Microsoft Teams for now (18/10/2017)
            // https://msdn.microsoft.com/en-us/microsoft-teams/bots#cards-and-buttons
            if (session.message.address.channelId = "msteams") {
                let link: CardAction = CardAction.openUrl(session, authorizationUrl, buttonLabel);

                signinCard = new ThumbnailCard(session)
                    .title("Authorization required!")
                    .text(actionLabel)
                    .buttons([link]);
            } else {
                signinCard = new SigninCard(session)
                    .text(actionLabel)
                    .button(buttonLabel, authorizationUrl);
            }

            let msg: Message = new Message(session).attachments([signinCard]);
            session.send(msg);
        } else {

            // if the user is  already signed-in, we check if the access token is expired
            let expiresOn: string = session.privateConversationData.expiresOn;
            let refreshToken: string = session.privateConversationData.refreshToken;

            if (new Date(expiresOn) >= new Date(Date.now())) {
                AuthentificationService.acquireTokenWithRefreshToken(refreshToken).then((response) => {
                    session.privateConversationData.accessToken = response.accessToken;
                    session.privateConversationData.expiresOn = response.expiresOn;
                    session.privateConversationData.refreshToken = response.refreshToken;

                    next();
                }).catch((errorMessage) => {
                    console.error(errorMessage);
                });
            } else {
                next();
            }
        }
    }

    /**
     * Check if we have adal config
     */
    private static checkAdal(): any {
        if (!this.adalConfig) {
            this.adalConfig = AuthentificationService.getAdalConfig();
        }
    }

    /**
     * Return Adal Config
     */
    public static getAdalConfig(): any {
        let adalConfig: any;
        adalConfig = {
            "clientId": process.env.AAD_CLIENT_ID,
            "clientSecret": process.env.AAD_CLIENT_SECRET,
            "authorityHostUrl": "https://login.microsoftonline.com/",
            "tenant": process.env.TENANT,
            "redirectUri": process.env.REDIRECT_URI,
            "resource": process.env.RESOURCE,
        };

        adalConfig.authorityUrl = adalConfig.authorityHostUrl +
            adalConfig.tenant;
        adalConfig.templateAuthzUrl = adalConfig.authorityUrl +
            "/oauth2/authorize?response_type=code&client_id=" +
            adalConfig.clientId +
            "&state=<state>&resource=" +
            adalConfig.resource +
            "&response_mode=form_post" +
            "&redirect_uri=" + adalConfig.redirectUri;

        return adalConfig;
    }
}