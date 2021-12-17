import {
  BotDeclaration,
  MessageExtensionDeclaration,
  PreventIframe,
} from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import {
  StatePropertyAccessor,
  CardFactory,
  TurnContext,
  MemoryStorage,
  ConversationState,
  ActivityTypes,
  TeamsActivityHandler,
  TaskModuleRequest,
  AttachmentLayoutTypes,
  TeamsInfo,
} from "botbuilder";
import HelpCard from "./dialogs/HelpDialog";
// import WelcomeCard from "./dialogs/WelcomeDialog";
 import WelcomeCard from "./WelcomeCard/welcomeCard";
 import TakeatourCard from "./take/Takeatour";
 import Takeatour from "./take/tour";
import Signincard from "./SigninCard/Signincard";
import {GetuserToken} from './ftoken';


 const log = debug("msteams");
 
 var storedata = [{}];
 var Sendername,Badgesent,image,comments,receiver;

 @BotDeclaration(
  "/api/messages",
  new MemoryStorage(),
  process.env.MICROSOFT_APP_ID,
  process.env.MICROSOFT_APP_PASSWORD
)
@PreventIframe("/thankZBot/aboutThankZ.html")
export class ThankZ extends TeamsActivityHandler {
  private readonly conversationState: ConversationState;
  private readonly dialogs: DialogSet;
  private dialogState: StatePropertyAccessor<DialogState>;
 
  /**
   * The constructor
   * @param conversationState
   */
  public constructor(conversationState, public adapter) {
    super();
    this.conversationState = conversationState;
    this.dialogState = conversationState.createProperty("dialogState");
    this.dialogs = new DialogSet(this.dialogState);
    // this.dialogs.add(new HelpDialog("help"));

    // Set up the Activity processing
    

    this.onMessage(
      async (context: TurnContext): Promise<void> => {
        //console.log(context);
        //storedata.push({context});
        // TODO: add your own bot logic in here
        switch (context.activity.type) {
          case ActivityTypes.Message:
            let text = TurnContext.removeRecipientMention(context.activity);
            text = text.toLowerCase();
            if (text.startsWith("hello")) {
              await context.sendActivities([
                { type: ActivityTypes.Typing },
                { type: "delay", value: 1000 }
              ]);
              await context.sendActivity("Oh, hello to you as well!");
              return;
            } else if (text.startsWith("help")) {
              await context.sendActivities([
                { type: ActivityTypes.Typing },
                { type: "delay", value: 1000 }
              ]);
              await context.sendActivity({ attachments: [HelpCard] });
              // await dc.beginDialog("help");
              return;
            }
             else if (text.startsWith("appreciate")) {
              await context.sendActivities([
                { type: ActivityTypes.Typing },
                { type: "delay", value: 1000 }
              ]);    
              //console.log(context.activity.conversation.tenantId,"Context");
              if(context.activity.conversation.tenantId === process.env.MY_TenantID) {
                await context.sendActivity({ attachments: [WelcomeCard] });
              } else {
                await context.sendActivity({ attachments: [HelpCard] });
              }
              return;
            } 
            else if (text.startsWith("take a tour")) {
              await context.sendActivities([
                { type: ActivityTypes.Typing },
                { type: "delay", value: 1000 }
              ]);          
              await context.sendActivity({ attachments: [ TakeatourCard, Takeatour], attachmentLayout: AttachmentLayoutTypes.Carousel });
              return;
            } 
            
            else {
              await context.sendActivity(
                `I am Sorry! I can't recognize what you are saying. Please try some other phrase like "help" to see what can I do for you`
              );
            }
            break;
          default:
            break;
        }
        // Save state changes
        return this.conversationState.saveChanges(context);
      }
    );

    this.onConversationUpdate(
      async (context: TurnContext): Promise<void> => {
        if (
          context.activity.membersAdded &&
          context.activity.membersAdded.length !== 0
        ) {
          for (const idx in context.activity.membersAdded) {
            if (
              context.activity.membersAdded[idx].id ===
              context.activity.recipient.id
            ) {
              await context.sendActivities([
                { type: ActivityTypes.Typing },
                { type: "delay", value: 1000 }
              ]);
              const member = await TeamsInfo.getMembers(context)
              let upn=member[0].userPrincipalName; 
              if(context.activity.conversation.tenantId === process.env.MY_TenantID) {
              let result =await GetuserToken(upn,{"receiverID":context.activity.from.id,"aadObjectId":context.activity.from.aadObjectId,"BotID":context.activity.recipient.id,"BotName":context.activity.recipient.name,"conversationType":context.activity.conversation.conversationType,"tenantId":context.activity.conversation.tenantId,"coversationid":context.activity.conversation.id,"Upnid":upn})
              }
              await context.sendActivity({ attachments: [Signincard] });
            
                
            }
          }
        }
      }
    );

    this.onMessageReaction(
      async (context: TurnContext): Promise<void> => {
        const added = context.activity.reactionsAdded;
        if (added && added[0]) {
          await context.sendActivity({
            textFormat: "xml",
            text: `That was an interesting reaction (<b>${added[0].type}</b>)`,
          });
        }
      }
    );
  }

 
  async handleTeamsTaskModuleFetch(
    context: TurnContext,
    taskModuleRequest: TaskModuleRequest
  ): Promise<any> {
    //console.log("hello");
    //console.log(taskModuleRequest);
    //console.log(context);
    //console.log("hai");
    //console.log(taskModuleRequest.data.status);

    // //console.log(localStorage.getItem('document'));
    
    const status = taskModuleRequest.data.status;
    if(status === "Take a tourb") {
        //console.log("Take a tour");
        await context.sendActivity({ attachments: [ TakeatourCard, Takeatour], attachmentLayout: AttachmentLayoutTypes.Carousel });
        return;
    }
    
    if (status === "Appreciate") {
      //console.log("Apppreciations");
      return {
        task: {
          type: "continue",
          value: {
            // card: this.getTaskModuleAdaptiveCard(),
            height: 600,
            width: 1000,
            title: "Sent Appreciation",
            url: "https://https://quadrathankz.azurewebsites.net/newTab",
          },
        },
      };
    }
    //    } else if (status === 'Badge') {
    //         return {
    //             task: {
    //                 type: 'continue',
    //                 value: {
    //                     // card: this.getTaskModuleAdaptiveCard(),
    //                     height: 550,
    //                     width: 900,
    //                     title: 'Adaptive Card: Inputs',
    //                     url: ' https://8123e66f1517.ngrok.io/newBadge'
    //                 },
    //             }
    //         };
    //    }
  }

  
  

  async handleTeamsTaskModuleSubmit(
    context: TurnContext,
    taskModuleRequest: TaskModuleRequest
  ): Promise<any> {
    
    let {name,Badges,Image,Receivers,managers,peers,comment} = taskModuleRequest.data;
    
    Sendername=name
    Badgesent=Badges
    image=Image
    comments=comment
    receiver=Receivers.header
   

    
      let address0 = {
        user:{
          id:context.activity.from.id,
          name: context.activity.from.name,
          aadObjectId:context.activity.from.aadObjectId
        },
        bot:
        {
            id: '28:a013c4a4-0683-4125-8c05-4004c2c3cc6f',
            name: 'Quadra Thankz'
        },
        conversation:
        {
            conversationType: 'personal',
            tenantId: '08948d7c-43ee-4cae-9f2c-67e0464345d8',
            id:context.activity.conversation.id
        },
        channelId: 'msteams',
        serviceUrl: 'https://smba.trafficmanager.net/apac/'
    };

    
    let address = {
      user:{
        id:Receivers.receiverID,
        name: Receivers.header,
        aadObjectId:Receivers.aadObjectId
      },
      bot:
      {
          id: Receivers.BotID,
          name: Receivers.BotName
      },
      conversation:
      {
          conversationType: Receivers.conversationType,
          tenantId: Receivers.tenantId,
          id:Receivers.coversationid   
      },
      channelId: 'msteams',
      serviceUrl: 'https://smba.trafficmanager.net/apac/'
  };

  //console.log(address)

    
    
//console.log(address0)

  let address1 = {
    user:{
      id:managers.receiverID,
      name: managers.header,
      aadObjectId:managers.aadObjectId
    },
    bot:
    {
        id:managers.BotID,
        name: managers.BotName
    },
    conversation:
    {
        conversationType: managers.conversationType,
        tenantId: managers.tenantId,
        id:managers.coversationid   
    },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/apac/'
};
//console.log(address1)

peers.map((val,key)=>{
  let address2 = {
    user:{
      id:val.receiverID,
      name: val.header,
      aadObjectId:val.aadObjectId
    },
    bot:
    {
        id:val.BotID,
        name: val.BotName
    },
    conversation:
    {
        conversationType: val.conversationType,
        tenantId: val.tenantId,
        id:val.coversationid   
    },
    channelId: 'msteams',
    serviceUrl: 'https://smba.trafficmanager.net/apac/'
};
  //console.log(address2)

  this.adapter.continueConversation(address2, async (turnContext) => {
    // If you encounter permission-related errors when sending this message, see
    await turnContext.sendActivity({ attachments: [this.getPeerCard()] });
  });

})

  
    this.adapter.continueConversation(address0, async (turnContext) => {
      // If you encounter permission-related errors when sending this message, see
      await turnContext.sendActivity({ attachments: [this.getselfCard()] });
    }); 

    this.adapter.continueConversation(address, async (turnContext) => {
      // If you encounter permission-related errors when sending this message, see
      await turnContext.sendActivity({ attachments: [this.getReceiveCard()] });
    });

    this.adapter.continueConversation(address1, async (turnContext) => {
      // If you encounter permission-related errors when sending this message, see
      await turnContext.sendActivity({ attachments: [this.getManagerCard()] });
    });

    
  // });  
  // }
  };

  


  getTaskModuleAdaptiveCard() {
    return CardFactory.adaptiveCard({
      version: "1.0.0",
      type: "AdaptiveCard",
      body: [
        {
          type: "TextBlock",
          text: "Select a badge",
          size: "large",
        },
        {
          type: "ColumnSet",
          columns: [
            {
              type: "Column",
              items: [
                {
                  type: "Image",
                  style: "Person",
                  url:
                    "http://curtforcouncil.com/wp-content/uploads/2018/05/thank-you.jpg",
                  size: "medium",
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
                  text: "ThankYou",
                  size: "large",
                  wrap: true,
                },
                {
                  type: "TextBlock",
                  spacing: "None",
                  text: "thankyou for ..............",
                  isSubtle: true,
                  wrap: true,
                },
              ],
              width: "stretch",
            },
          ],
        },
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Next",
        },
      ],
    });
  }


  getReceiveCard(){
    return CardFactory.adaptiveCard({
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
                              "text": `${Sendername} has send you !`,
                              "wrap": true
                          },
                          {
                              "type": "TextBlock",
                              "spacing": "None",
                              "weight": "Bolder",
                              "size": "Large",
                              "text": `${Badgesent}`,
                              "horizontalAlignment": "Center",
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
                  "width": "150px",
                  
                  
              }]
          },
          {
              "type": "TextBlock",
              "text": `Hi ${receiver}!`,
              "wrap": true
          },
          {
              "type": "TextBlock",
              "text": `${comments}`,
              "wrap": true
          }
      ]
      
  })
  }

  // manager

  getManagerCard(){
    return CardFactory.adaptiveCard({
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
                              "text": `${Sendername} has send an card to ${receiver}!`,
                              "wrap": true
                          },
                          {
                            "type": "TextBlock",
                            "spacing": "None",
                            "weight": "Bolder",
                            "size": "Large",
                            "text": `${Badgesent}`,
                            "isSubtle": true,
                            "horizontalAlignment": "Center",
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
                  "width": "150px"
              }]
          },
          {
              "type": "TextBlock",
              "text": `Hi ${receiver}`,
              "wrap": true
          },
          {
              "type": "TextBlock",
              "text": `${comments}`,
              "wrap": true
          }
      ],
      
  })
  }


  // peer

  getPeerCard(){
    return CardFactory.adaptiveCard({
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
                              "text": `${Sendername} has send card to  ${receiver}`,
                              "wrap": true
                          },
                          {
                            "type": "TextBlock",
                            "spacing": "None",
                            "weight": "Bolder",
                            "size": "Large",
                            "text": `${Badgesent}`,
                            "horizontalAlignment": "Center",
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
                  "width": "150px"
              }]
          },
          {
            "type": "TextBlock",
            "text": `Hi ${receiver}`,
            "wrap": true
        },
        {
            "type": "TextBlock",
            "text": `${comments}`,
            "wrap": true
        }
      ],
      
  })
  }
 
  getselfCard(){
    return CardFactory.adaptiveCard({
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
                              "text": `You have sent card to  ${receiver}`,
                              "wrap": true
                          },
                          {
                            "type": "TextBlock",
                            "spacing": "None",
                            "weight": "Bolder",
                            "size": "Large",
                            "text": `${Badgesent}`,
                            "horizontalAlignment": "Center",
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
                  "width": "150px"
              }]
          },
          {
            "type": "TextBlock",
            "text": `Hi ${receiver}`,
            "wrap": true
        },
        {
            "type": "TextBlock",
            "text": `${comments}`,
            "wrap": true
        }
      ],
      
  })
  }

}



