import { AuthentificationService } from "./authentification-service";
import * as fetch from "node-fetch";

require("dotenv").config();
export class SharepointService {

    private static adalConfig: any;

    /**
     * Launch a search query in SharePoint
     */
    public static doSearch(query: string, accessToken: string): Promise<any> {
        let adalConfig: any = AuthentificationService.getAdalConfig();
        let p: Promise<any> = new Promise((resolve, reject) => {
            let endpointUrl: string = adalConfig.resource + process.env.SEARCH_PATH + "querytext='" + query + "'";
            fetch(endpointUrl, {
                method: "GET",
                headers: {
                    "Authorization": "Bearer " + accessToken,
                    "Accept": "application/json;odata=verbose"
                }
            }).then((res) => {
                return res.json();
            }).then((json) => {
                resolve(json);
            }).catch((err) => {
                reject(err);
            });
        });

        return p;
    }
}