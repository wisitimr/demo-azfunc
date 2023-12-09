const { app } = require('@azure/functions');
const axios = require("axios");

app.storageBlob('testLinkStorageBlobTrigger', {
    path: 'cms/test-link/{name}.json',
    connection: 'AzureWebJobsStorage',
    handler: async (blob, context) => {
        context.log(
            `Storage blob function processed blob "${context.triggerMetadata.name}" with size ${blob.length} bytes`
        );
        let blob_input_product = null;
        try {
            blob_input_product = JSON.parse(blob.toString());
        } catch (error) {
            console.log(error);
        }

        // if we can parse input file as json
        if (blob_input_product && blob_input_product.length) {
            let date = new Date(),
                dateString = date.toLocaleString("en-US", {
                    timeZone: "Asia/Bangkok",
                }),
                datalist = new Array(),
                attachlist = new Array(),
                callTests = new Array(),
                attachments = new Array();

            for (const [i, item] of blob_input_product.entries()) {
                const attachment = {
                    contentType: "application/vnd.microsoft.card.adaptive",
                    content: {
                        type: "AdaptiveCard",
                        msteams: {
                            width: "Full"
                        },
                        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                        version: "1.6",
                        body: [
                            {
                                type: "ColumnSet",
                                columns: [
                                    {
                                        type: "Column",
                                        items: [
                                            {
                                                type: "TextBlock",
                                                text: "MyAIS - link monitor alert",
                                                size: "large",
                                                weight: "bolder",
                                                wrap: true
                                            },
                                            {
                                                type: "TextBlock",
                                                spacing: "None",
                                                text: "Alert on " + dateString,
                                                isSubtle: true,
                                                wrap: true
                                            }
                                        ],
                                        width: "stretch"
                                    }
                                ]
                            },
                            {
                                type: "ColumnSet",
                                columns: [
                                    {
                                        type: "Column",
                                        items: [
                                            {
                                                type: "Image",
                                                url: process.env.IMAGE_URL,
                                                size: "Small",
                                            },
                                        ],
                                        width: "auto",
                                    },
                                    {
                                        type: "Column",
                                        items: [
                                            {
                                                type: "TextBlock",
                                                weight: "Bolder",
                                                text: " ",
                                                wrap: true,
                                            }
                                        ],
                                        width: "stretch",
                                    },
                                ],
                            }
                        ]
                    }
                }
                if (item.topic) {
                    attachment.content.body[1].columns[1].items[0].text = item.topic
                }
                if (item.description) {
                    attachment.content.body[1].columns[1].items[1] = {
                        type: "TextBlock",
                        spacing: "None",
                        text: item.description,
                        isSubtle: true,
                        wrap: true,
                    }
                }
                for (const iterator of item.data) {
                    datalist.push(iterator);
                    attachlist.push(attachment);
                    callTests.push(invalidLink(context, iterator.url));
                    // const error = await invalidLink(context, iterator.url);
                    // if (error && error.message) {
                    //     invalidDetial.push({
                    //         type: "ColumnSet",
                    //         columns: [
                    //             {
                    //                 type: "Column",
                    //                 width: "100px",
                    //                 items: [
                    //                     {
                    //                         type: "TextBlock",
                    //                         text: iterator.id,
                    //                         weight: "bolder",
                    //                     },
                    //                 ],
                    //             },
                    //             {
                    //                 type: "Column",
                    //                 width: "auto",
                    //                 items: [
                    //                     {
                    //                         type: "TextBlock",
                    //                         text: `[${iterator.url}](${iterator.url})`,
                    //                         wrap: true,
                    //                     }
                    //                 ]
                    //             },
                    //             {
                    //                 type: "Column",
                    //                 width: "auto",
                    //                 items: [
                    //                     {
                    //                         type: "TextBlock",
                    //                         text: `(error: ${error.message})`,
                    //                         wrap: true,
                    //                     }
                    //                 ]
                    //             }
                    //         ],
                    //     });
                    // }
                    if (callTests.length == Number(process.env.TEST_CONCURRENT) || blob_input_product.length == i + 1) {
                        const errors = await Promise.all(callTests);
                        for (const [i, error] of errors.entries()) {
                            if (error && error.message) {
                                attachlist[i].content.body.push({
                                    type: "ColumnSet",
                                    columns: [
                                        {
                                            type: "Column",
                                            width: "100px",
                                            items: [
                                                {
                                                    type: "TextBlock",
                                                    text: datalist[i].id,
                                                    weight: "bolder",
                                                },
                                            ],
                                        },
                                        {
                                            type: "Column",
                                            width: "auto",
                                            items: [
                                                {
                                                    type: "TextBlock",
                                                    text: `[${datalist[i].url}](${datalist[i].url})`,
                                                    wrap: true,
                                                }
                                            ]
                                        },
                                        {
                                            type: "Column",
                                            width: "auto",
                                            items: [
                                                {
                                                    type: "TextBlock",
                                                    text: `(error: ${error.message})`,
                                                    wrap: true,
                                                }
                                            ]
                                        }
                                    ],
                                });
                                attachments.push([attachlist[i]]);
                            }
                        }
                        datalist = new Array();
                        attachlist = new Array();
                        callTests = new Array();
                    }
                }
            }
            let notifications = new Array();
            context.log('Total sendTeamsNotification ========= ', attachments.length)
            if (attachments.length) {
                for (const [i, attach] of attachments.entries()) {
                    notifications.push(sendTeamsNotification(context, attach));
                    if (notifications.length == Number(process.env.WEBHOOK_CONCURRENT) || attachments.length == i + 1) {
                        await Promise.all(notifications);
                        notifications = new Array();
                    }
                }
            }
        }
    }
});

async function invalidLink(context, link) {
    try {
        await axios.get(link, { timeout: Number(process.env.TEST_TIMEOUT) });
    } catch (error) {
        if (error && error.message) {
            context.log("invalidLink Error: " + error.message);
            return {
                message: error.message
            }
        }
    }
    return null;
}

async function sendTeamsNotification(context, attachments) {
    try {
        let url = process.env.WEBHOOK_URL
        if (context.triggerMetadata.name.startsWith("dev/")) {
            url = process.env.DEV_WEBHOOK_URL;
        } else if (context.triggerMetadata.name.startsWith("sit/")) {
            url = process.env.SIT_WEBHOOK_URL;
        } else if (context.triggerMetadata.name.startsWith("uat/")) {
            url = process.env.UAT_WEBHOOK_URL;
        }
        console.log(`AXIOS: POST TO ${url}`);
        await axios.post(url, {
            type: "message",
            attachments,
        });
    } catch (error) {
        context.log("sendTeamsNotification Error: " + error.message);
        throw error;
    }
}

