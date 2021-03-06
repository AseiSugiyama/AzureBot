/*jshint esversion: 6 */
require('dotenv').config();

const restify = require('restify');
const builder = require('botbuilder');
const ticketsApi = require('./ticketsApi');
const fs = require('fs');

const listenPort = process.env.port || process.env.PORT || 3978;
const ticketSubmissionUrl = process.env.TICKET_SUBMISSION_URL || `http://localhost:${listenPort}`;

// Setup Restify Server
var server = restify.createServer();
server.listen(listenPort, '::', () => {
  console.log('Server Up');
} );

// Setup body parser and tickets api
server.use(restify.bodyParser());
server.post('/api/tickets', ticketsApi);

// Azure Search API Client
const azureSearch = require('./azureSearchApiClient');

const azureSearchQuery = azureSearch({
    searchName: process.env.AZURE_SEARCH_ACCOUNT,
    indexName: process.env.AZURE_SEARCH_INDEX,
    searchKey: process.env.AZURE_SEARCH_KEY
});


// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Listen for messages from users
server.post('/api/messages', connector.listen());

// Receive messages from the user and respond by echoing each message back (prefixed with 'You said:')
var bot = new builder.UniversalBot(connector,
  /* Dialogs */
  (session) => {
    session.endDialog(`I'm sorry, I did not understand '${session.message.text}'. Type 'help' to know more about me :)`);
  }
);

var luisRecognizer = new builder.LuisRecognizer(process.env.LUIS_MODEL_URL).onEnabled(function (context, callback) {
  var enabled = context.dialogStack().length === 0;
  callback(null, enabled);
});
bot.recognizer(luisRecognizer);

bot.dialog('SubmitTicket',
  // Dialogs
  [
    // 1st
    (session, args, next) => {
      var category = builder.EntityRecognizer.findEntity(args.intent.entities, 'category');
      var severity = builder.EntityRecognizer.findEntity(args.intent.entities, 'severity');

      if (category && category.resolution.values.length > 0) {
        session.dialogData.category = category.resolution.values[0];
      }

      if (severity && severity.resolution.values.length > 0) {
        session.dialogData.severity = severity.resolution.values[0];
      }

      session.dialogData.description = session.message.text;

      if (!session.dialogData.severity) {
        var choices = ['high', 'normal', 'low'];
        builder.Prompts.choice(session,
           'which is the severity of this problem?',
            choices,
            { listStyle : builder.ListStyle.button }
          );
      }else{
        next();
      }
    },
    // 2nd
    (session, result, next) => {
      if (!session.dialogData.severity) {
        session.dialogData.severity = result.response.entity;
      }

      if (!session.dialogData.category) {
        builder.Prompts.text(session, 'Which would be the category for this ticket (software, hardware, network, and so on)?');
      }else{
        next();
      }
    },
    // 4th
    (session, result, next) => {
      if (!session.dialogData.category) {
        session.dialogData.category = result.response;
      }

      var message = `Great! I'm going to create a "${session.dialogData.severity}" severity ticket in the "${session.dialogData.category}" category. ` +
      `The description I will use is "${session.dialogData.description}". Can you please confirm that this information is correct?`;

      // next result.response will return as a boolean value (true or false).
      builder.Prompts.confirm(
        session,
        message,
        {
          listStyle : builder.ListStyle.button
        }
      );
    },
    (session, result, next) => {
      if (result.response){
        var data = {
          category: session.dialogData.category,
          severity: session.dialogData.severity,
          description: session.dialogData.description
        };

        const client = restify.createJsonClient({ url: ticketSubmissionUrl });

        client.post('/api/tickets', data, (err, request, response, ticketId) => {
          if(err || ticketId == -1){
            session.send('Something went wrong while I was saving your ticket. Please try again later.');

          }else{
            session.send(`Awesome! Your ticket has been created with the number ${ticketId}.`);
            session.send(new builder.Message(session).addAttachment({
              contentType: "application/vnd.microsoft.card.adaptive",
              content: createCard(ticketId, data)
            }));
          }
        });

        session.endDialog();
      }else{
        session.endDialog('Ok. The ticket was not created. You can start again if you want.');
      }
    }
  ]
)
.triggerAction({
  matches: 'SubmitTicket'
});

bot.dialog('ExploreKnowledgeBase', [
  (session, args) => {
    var category = builder.EntityRecognizer.findEntity(args.intent.entities, 'category');

    if (!category) {
      return session.endDialog('Try typing something like _explore hardware_.');
    }

    //search by category
    azureSearchQuery('fileter=' + encodeURIComponent(`category eq ${category.entity}`), (error, result) => {
      if (error) {
        console.log(error);
        session.endDialog('Ooops! Something went wrong while contacting Azure Search. Please try again later.');
      } else {
        var msg = `These are some articles I\'ve found in the knowledge base for the _'${category.entity}'_ category:`;
        result.value.forEach( (article) => {
          msg += `\n * ${article.title}`;
        });
        session.endDialog(msg);
      }
    }); 
  }
]).triggerAction({
  matches:'ExploreKnowledgeBase'
});


bot.dialog('Help',
  (session, args, next) => {
    session.endDialog(`I'm the help desk bot and I can help you create a ticket.\n` +
                      `You can tell me things like _I need to reset my password_ or _I cannot print_.`);
  }
).triggerAction({
  matches: 'Help'
});


// EOB
/**
  * createCard
  * @param ticketId : integer
  * @param data : Object
  */

const createCard = (ticketId, data) => {
  var cardText = fs.readFileSync('./cards/ticket.json', 'UTF-8');

  cardText = cardText.replace(/{ticketId}/g, ticketId)
                    .replace(/{severity}/g, data.severity)
                    .replace(/{category}/g, data.category)
                    .replace(/{description}/g, data.description);
  return JSON.parse(cardText);
};
