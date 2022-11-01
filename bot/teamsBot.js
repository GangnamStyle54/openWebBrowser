const axios = require("axios");
const querystring = require("querystring");
const open = require("open");
const { TeamsActivityHandler, CardFactory, TurnContext, ActionTypes, MessageFactory } = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const cardTools = require("@microsoft/adaptivecards-tools");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // record the likeCount
    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "google": {
          open("https://google.com/");
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(context, invokeValue) {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200 };
    }
  }

  // Message extension Code
  // Action.
  async handleTeamsMessagingExtensionSubmitAction(context, action) {
    switch (action.commandId) {
      case "postComment":
        cardActivityAsync(context, action);
        break;
      case "openWebUrl":
        await openWebUrlCommand(context);
        break;
      default:
        throw new Error("NotImplemented");
    }
  }

  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const heroCard = CardFactory.heroCard(obj.package.name);
      const preview = CardFactory.heroCard(obj.package.name);
      preview.content.tap = {
        type: "invoke",
        value: { name: obj.package.name, description: obj.package.description },
      };
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  async handleTeamsMessagingExtensionSelectItem(context, obj) {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }

  // Link Unfurling.
  handleTeamsAppBasedLinkQuery(context, query) {
    const attachment = CardFactory.thumbnailCard("Thumbnail Card", query.url, [query.url]);

    const result = {
      attachmentLayout: "list",
      type: "result",
      attachments: [attachment],
    };

    const response = {
      composeExtension: result,
    };
    return response;
  }
}

async function cardActivityAsync(context, action) {
  const cardActions = [
    {
        type: ActionTypes.MessageBack,
        title: 'Message all members',
        value: null,
        text: 'MessageAllMembers'
    },
    {
        type: ActionTypes.MessageBack,
        title: 'Who am I?',
        value: null,
        text: 'whoami'
    },
    {
        type: ActionTypes.MessageBack,
        title: 'Delete card',
        value: null,
        text: 'Delete'
    }
];

    await sendUpdateCard(context, cardActions);
}

async function openWebUrlCommand(context) {
  //Open web url
  await open('https://google.pt?query=test');
}

async function sendUpdateCard(context, cardActions) {
  const data = context.activity.value;
  const adaptiveCard = data.messagePayload.attachments[0];

  const adaptiveCardContent = JSON.parse(adaptiveCard.content);

  data.count += 1;
  cardActions.push({
      type: ActionTypes.MessageBack,
      title: 'Update Card',
      value: data,
      text: 'UpdateCardAction'
  });
  const card = CardFactory.heroCard(
      'Updated card',
      `Update count: ${ data.count }`,
      null,
      cardActions
  );
  card.id = adaptiveCard.id;
  const message = MessageFactory.attachment(card);
  message.id = adaptiveCard.id;

  try {
    await context.updateActivity(message);
    console.log("não está a dar erro");
  } catch (error) {
    console.log(error);
  }
}

module.exports.TeamsBot = TeamsBot;
