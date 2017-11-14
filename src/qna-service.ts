import { unescape } from "querystring";
import { SharepointService } from "./sharepoint-service";
import { QnAMakerRecognizer, QnAMakerTools, QnAMakerDialog, IQnAMakerResults } from "botbuilder-cognitiveservices";
import * as builder from "botbuilder";
import * as decode from "unescape";

require("dotenv").config();

export class QnAService {
    private static initialized: boolean;
    private static qnAMakerTools: any;
    private static customQnAMakerTools: CustomQnAMakerTools;

    private static qnARecognizer: any;

    /**
     * Initialize the service
     */
    private static initialize(bot: builder.UniversalBot): any {
        this.qnARecognizer = new QnAMakerRecognizer({
            knowledgeBaseId: process.env.QNA_KNOWLEDGEBASE_ID,
            subscriptionKey: process.env.QNA_SUBSCRIPTION_KEY,
            top: 3
        });

        this.qnAMakerTools = new QnAMakerTools();
        bot.library(this.qnAMakerTools.createLibrary());

        this.customQnAMakerTools = new CustomQnAMakerTools();
        bot.library(this.customQnAMakerTools.createLibrary());
    }

    /**
     * Get QnaMakerDialog
     */
    public static basicQnAMakerDialog(bot: builder.UniversalBot): any {
        if (!this.initialized) {
            this.initialize(bot);
        }

        let qnaMakerDialog: any = new QnAMakerDialog({
            recognizers: [this.qnARecognizer],
            defaultMessage: "No match! Try changing the query terms!",
            qnaThreshold: 0.3,
            feedbackLib: this.customQnAMakerTools
        });

        qnaMakerDialog.respondFromQnAMakerResult = (session, qnaMakerResult) => {
            processAnswer(session, qnaMakerResult.answers[0].answer, null);
        };

        return qnaMakerDialog;
    }
}


export class CustomQnAMakerTools {
    private lib: builder.Library;

    constructor() {
        this.lib = new builder.Library("customQnAMakerTools");
        this.lib.dialog("answerSelection", [
            function (session: builder.Session, args: any): any {
                let qnaMakerResult: any = args;
                session.dialogData.qnaMakerResult = qnaMakerResult;
                let questionOptions: any[] = [];

                qnaMakerResult.answers.forEach((qna) => {
                    questionOptions.push(decode(qna.questions[0]));
                });

                questionOptions.push("Aucun de ces résultats");

                let promptOptions: any = {
                    listStyle: builder.ListStyle.button
                };

                builder.Prompts.choice(session, "J'ai trouvé plusieurs réponse qui pourrait vous convenir, choisissez l'une d'elle:"
                    , questionOptions, promptOptions);

            },
            function (session: builder.Session, results: any): any {
                let qnaMakerResult: any = session.dialogData.qnaMakerResult;
                let filteredResult: any = qnaMakerResult.answers.filter(
                    (qna) => {
                        return decode(qna.questions[0]) === results.response.entity;
                    }
                );

                let selectedQnA: any = filteredResult[0];

                if (selectedQnA) {
                    processAnswer(session, selectedQnA.answer, selectedQnA);
                } else {
                    session.endDialog();
                }
            }]);
    }

    public createLibrary(): builder.Library {
        return this.lib;
    }

    public answerSelector(session: builder.Session, options: IQnAMakerResults): void {
        session.beginDialog("customQnAMakerTools:answerSelection", options || {});
    }
}

/**
 * Check a QnA Answer and if it follow the pattern {query} do a search on Sharepoint according the request.
 */
function processAnswer(session: builder.Session, answer: any, selectedQnA: any): any {
    let searchQuery: string = answer.match(/^\{(.*?)\}$/);

    if (searchQuery) {
        let accessToken: string = session.privateConversationData.accessToken;

        SharepointService.doSearch(searchQuery[1], accessToken).then((res) => {
            if (res.error) {
                session.send("Error: %s", res.error.message.value);
            } else {
                let cards: builder.ThumbnailCard[] = new Array<builder.ThumbnailCard>();
                let results: any = res.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;

                if (results.length > 0) {
                    results.forEach((value) => {
                        let valueRes: Array<any> = value.Cells.results;

                        let title: string = valueRes.find(o => o.Key === "Title").Value;
                        let link: any = builder.CardAction.openUrl(session, valueRes.find(o => o.Key === "Path").Value, "View");
                        let fileType: string = valueRes.find(o => o.Key === "FileType").Value;
                        let hitHighlightedSummary: string = valueRes.find(o => o.Key === "HitHighlightedSummary").Value;
                        hitHighlightedSummary = hitHighlightedSummary.replace(/<c0>|<\/c0>/g, "").replace(/<ddd\/>/g, "");
                        let elt: builder.ThumbnailCard = new builder.ThumbnailCard(session).title(title)
                            .text(unescape(hitHighlightedSummary)).subtitle("Type: " + fileType).buttons([link]);

                        cards.push(elt);
                    });

                    let reply: any = new builder.Message(session)
                        .attachmentLayout(builder.AttachmentLayout.carousel)
                        .attachments(cards);

                    console.log(reply);
                    session.send(reply);

                    session.endDialogWithResult(selectedQnA);

                } else {
                    session.send("Sorry, we didn't find anything for '\%s\'", searchQuery.toString());
                    if (selectedQnA) {
                        session.endDialogWithResult(selectedQnA);
                    } else {
                        session.endDialog();
                    }
                }
            }
        }).catch((error) => {
            session.send(error);
            if (selectedQnA) {
                session.endDialogWithResult(selectedQnA);
            } else {
                session.endDialog();
            }
        });
    } else {
        session.send(answer);

        if (selectedQnA) {
            session.endDialogWithResult(selectedQnA);
        } else {
            session.endDialog();
        }
    }
}