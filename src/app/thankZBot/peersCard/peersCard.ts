import { CardFactory } from "botbuilder";
import image from './peersCardImage';

const peersCard = CardFactory.adaptiveCard({
    version: "1.2",
    "type": "AdaptiveCard",
    "body": [
        {
            "type": "ColumnSet",
            "columns": [
                {
                    "type": "Column",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Suganthi sent thank you apprication to",
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "spacing": "None",
                            "weight": "Bolder",
                            "size": "Large",
                            "text": "Pavithrra Sekar",
                            "isSubtle": true,
                            "wrap": true
                        }
                    ],
                    "width": "stretch"
                },
            ],
        },
        {
            "type": "Container",
            "items": [{
                "type": "Image",
                "horizontalAlignment": "center",
                "url": `${image}`,
            }]
        },
        {
            "type": "TextBlock",
            "text": "Hi Pavithrra Sekar!",
            "wrap": true
        },
        {
            "type": "TextBlock",
            "text": "Its been two months since we have deliverd the project and this has been a great learing experience for me. Thank you for help!",
            "wrap": true
        }
    ],
    "actions": [
        {
            "type": "Action.ShowCard",
            "title": "Add comment",
            "card": {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "Input.Text",
                        "id": "comment",
                        "placeholder": "Add a comment",
                        "isMultiline": true
                    }
                ],
                "actions": [{
                    "type": "Action.Submit",
                    "title": "Submit",
                    "data": {
                        "Status": "Submit",
                    }
                }
                ],
            }
        },
        {
            "type": "Action.Submit",
            "title": "Chat",
            "data": {
                "msteams": {
                    "type": "imBack",
                    "value": "Chat"
                }
            }
        },
        {
            "type": "Action.Submit",
            "title": "Like",
            "data": {
                "msteams": {
                    "type": "imBack",
                    "value": "Like"
                }
            }
        }
    ]
})
export default peersCard;