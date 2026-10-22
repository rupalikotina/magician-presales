const {
  TeamsActivityHandler,
  MessageFactory,
  InputHints,
  CardFactory,
  TeamsInfo
} = require('botbuilder');
const {
  AdaptiveCards
} = require("@microsoft/adaptivecards-tools");
const {
  InvokeResponseFactory
} = require("@microsoft/teamsfx");
const { postLoops } = require('../internal/loopsClient');
const { logger, elapsedMs } = require('../internal/logger');
const { sendTeamsActivity, updateTeamsActivity, logConnectorError } = require('../internal/teamsResponse');
const { noteActivity } = require('../internal/duplicateDetector');

// Constants
const BACKEND_URL = 'https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events';

class LoopsCommandHandler {
  constructor() {
      // Regex pattern to detect and capture the 'loops' command followed by any arguments.
      this.triggerPatterns = /^(loops|ask)(\s+(.*?))?$/i;
  }


  /**
   * Handles the reception of a "loops" command.
   */
  async handleCommandReceived(context, message) {
    // Idempotency: suppress duplicate invocations for the same activity.
    // A missing activity id is not a key, so unrelated commands are not collapsed.
      try {
          const observation = noteActivity(context && context.activity, {
              suppress: true,
              text: message && message.text,
          });
          if (observation.suppress) {
              return;
          }
      } catch (e) {
          logger.error("Duplicate guard error", {
              event: "TEAMS_ACTIVITY_FAILED",
              error: e,
          });
      }
      
      logger.info("App received command", {
          event: "TEAMS_ACTIVITY_RECEIVED",
          activityType: context.activity && context.activity.type,
          text: message.text,
          activityId: context.activity && context.activity.id,
      });

      // Extract the argument text after the "loops" command.
      const command = message.matches ? message.matches[1].trim() : '';
      const commandText = (message.matches && message.matches[3]) ? message.matches[3].trim() : '';
      logger.info("Activity classified", {
          event: "TEAMS_ACTIVITY_CLASSIFIED",
          actionType: command || "loops",
          text: message.text,
          activityId: context.activity && context.activity.id,
      });

      const actionType = 'FetchInputAdaptiveCard';
      const formData = {};

      // Construct the payload for the backend with the extracted command details.
      const payload = {
          ...await constructPayloadForBackend(actionType, formData, context),
          command: command,
          argument: commandText,
          threadId: extractThreadId(context.activity.conversation.id),
          eventTimestamp: context.activity.timestamp,
          eventId: context.activity.id,
          from: context.activity.from,
          to: context.activity.recipient,
      };
      const payloadJsonString = JSON.stringify(payload);

      try {
          let response;
          try {
              response = await postLoops(BACKEND_URL, payloadJsonString);
          } catch (backendError) {
              logger.error("Error with the backend request", {
                  event: "TEAMS_ACTIVITY_FAILED",
                  durationMs: elapsedMs(),
                  error: backendError,
                  response: backendError && backendError.response && backendError.response.data,
              });
              const cardJson = getDefaultAdaptiveCard(backendError);
              const errorMessageActivity = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
              await sendTeamsActivity(context, errorMessageActivity);
              throw new Error("Backend request failed.");
          }
          // Extract or default to error message based on backend response.
          let adaptiveCardJson = '';
          let hiddenData = response.data.hiddenData;
          if (!response.data.cardJson) {
              const actionType = 'ProcessAdaptiveCardForm';
              logger.info("Activity classified", {
                  event: "TEAMS_ACTIVITY_CLASSIFIED",
                  actionType,
                  activityId: context.activity && context.activity.id,
              });
              // Create the payload
              const payload = {
                  ...await constructPayloadForBackend(actionType, formData, context),
                  from: context.activity.from,
                  to: context.activity.recipient,
                  threadId: extractThreadId(context.activity.conversation.id),
                  eventTimestamp: context.activity.timestamp,
                  eventId: context.activity.id,
                  hiddenData: hiddenData
              };

              // Default 'isButton' and 'responseAsReply' to false if not present
              // Extract buttonData
              const buttonData = formData.buttonData || {};

              // Default 'isButton' and 'responseAsReply' to false if not present
              const isButton = buttonData.isButton || false;
              const responseAsReply = buttonData.responseAsReply || false;

              const payloadJsonString = JSON.stringify(payload);

              const response = await postLoops(BACKEND_URL, payloadJsonString);
              await processBackendResponse(response, context, responseAsReply, null);
          } else {
              adaptiveCardJson = JSON.parse(response.data.cardJson);
          }
          // Construct the adaptive card attachment from the JSON.
          const adaptiveCardAttachment = {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: adaptiveCardJson
          };

          // Update the processing activity with the final response card.
          if (adaptiveCardJson) {
              await sendTeamsActivity(context, {
                  attachments: [adaptiveCardAttachment]
              });
          }
      } catch (error) {
          logger.error("Error handling the '" + command + "' command", {
              event: "TEAMS_ACTIVITY_FAILED",
              command,
              error,
          });

          const errorMessage = 'An error occurred while processing your command.';
          const cardJson = getDefaultAdaptiveCard(errorMessage);
          const errorMessageActivity = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));

          await sendTeamsActivity(context, errorMessageActivity);
      }
  }


  /**
   * Handles invoke activities, typically from actionable messages or adaptive cards.
   */
  async onInvoke(context) {
      const {
          activity
      } = context;

      try {
          noteActivity(activity, { suppress: false });
      } catch (error) {
          logger.error("Duplicate observation failed", {
              event: "TEAMS_ACTIVITY_FAILED",
              error,
          });
      }

      logger.info("Command invoke received", {
          event: "TEAMS_ACTIVITY_RECEIVED",
          activityType: activity && activity.type,
          activityName: activity && activity.name,
          value: activity && activity.value,
          activityId: activity && activity.id,
      });
      logger.info("Activity classified", {
          event: "TEAMS_ACTIVITY_CLASSIFIED",
          actionType: activity && activity.name,
          value: activity && activity.value,
          activityId: activity && activity.id,
      });

      // Case when the user opens a task module.
      if (activity.name === 'task/fetch') {
          const cardJson = getDefaultAdaptiveCard('Provide the required input.');
          const taskModuleResponse = {
              task: {
                  type: 'continue',
                  value: {
                      card: CardFactory.adaptiveCard(cardJson),
                      height: 400,
                      width: 500,
                      title: "Response Card"
                  }
              }
          };
          return InvokeResponseFactory.createInvokeResponse(taskModuleResponse);
      }
      // Case when the user submits data from a task module.
      else if (activity.name === 'task/submit') {
          const submittedData = activity.value;

          // Placeholder: Handle and process the submitted data as required.

          const thankYouCard = {
              type: "AdaptiveCard",
              version: "1.4",
              body: [{
                  type: "TextBlock",
                  size: "Medium",
                  weight: "Bolder",
                  text: "Thank you for submitting!"
              }],
              $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
          };
          const thankYouActivity = MessageFactory.attachment(CardFactory.adaptiveCard(thankYouCard));
          await sendTeamsActivity(context, thankYouActivity);
      }
      // Default case for unknown invoke actions.
      else {
          const errorMessage = 'Unknown or unsupported invoke action.';
          const cardJson = getDefaultAdaptiveCard(errorMessage);
          const errorMessageActivity = MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
          await sendTeamsActivity(context, errorMessageActivity);
      }
  }
}



/**
* Constructs a payload to be sent to the backend containing 
*/
async function constructPayloadForBackend(actionType, data, context) {
  // Extract channel and team details from the context.
  const channelId = context.activity.channelData && context.activity.channelData.channel ? context.activity.channelData.channel.id : 'unknown';
  let teamId = context.activity.channelData.team ? context.activity.channelData.team.id : undefined;
  let teamName = context.activity.channelData.team ? context.activity.channelData.team.name : "a channel";

  if (teamId) {
      try {
          // Fetch team details using the teamId.
          const teamDetails = await TeamsInfo.getTeamDetails(context, teamId);
          teamName = teamDetails.name;
          teamId = teamDetails.aadGroupId;
      } catch (err) {
          logConnectorError("Error fetching team details", err);
          teamName = "Unknown"; // Or any other default value.
      }
  } else {
      teamName = "Unknown"; // Default value if no team data is available.
  }

  // Return the constructed payload.
  return {
      action: actionType,
      conversationId: channelId,
      channelId: channelId,
      teamsTenantId: context.activity.channelData.tenant.id,
      teamId: teamId,
      teamName: teamName,
      data: data,
      rawEvent: context
  };
}



/**
* Generates and returns an adaptive card with a message.
* This card is meant to be used as a fallback when the system is unable
* to retrieve or render the primary form/data.
*/
function getDefaultAdaptiveCard(errorMessage) {
  return {
      type: "AdaptiveCard",
      version: "1.4",
      body: [{
              type: "TextBlock",
              size: "Medium",
              weight: "Bolder",
              text: "Error"
          },
          {
              type: "TextBlock",
              text: errorMessage || "An error occurred while processing your command.",
              wrap: true
          }
      ],
      $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
  };
}



/**
* Generates a unique string ID.
*
* The generated ID combines the current time in milliseconds (since 1970) 
* and a random string, both represented in base 36, to ensure uniqueness.
*/
function generateUniqueId() {
  // Get the current time in milliseconds since 1970 and convert it to a base 36 string.
  const timestampPart = Date.now().toString(36);

  // Generate a random number, convert it to a base 36 string, and take a substring of it.
  const randomPart = Math.random().toString(36).substr(2, 9);

  // Concatenate the timestamp part and the random part to form the unique ID.
  return timestampPart + randomPart;
}


function loadCommandsFromManifest() {
  try {
      // Directly require the manifest file as a JSON module
      const manifest = require("../../appPackage/manifest.json");

      if (!manifest || !manifest.bots || !Array.isArray(manifest.bots)) {
          throw new Error('Manifest does not contain "bots" array.');
      }

      // Assuming you want to extract commands from the first bot in the array
      const botCommands = manifest.bots[0].commandLists.flatMap(commandList => commandList.commands);

      if (!Array.isArray(botCommands)) {
          throw new Error('Bot commands are not structured as expected.');
      }

      return botCommands.map(cmd => cmd.title);
  } catch (error) {
      logger.error("Error loading commands from manifest", {
          event: "TEAMS_ACTIVITY_FAILED",
          error,
      });
      return []; // Return an empty array or handle the error as needed
  }
}



/**
* Process the response from the backend and send the appropriate message/card 
* to the user based on the response content.
*/
async function processBackendResponse(response, context, responseAsReply, activityId) {
  try {
      // Check if backend response indicates successful processing.
      if (response.data.status === "ok") {
          logger.info("Backend response", {
              event: "LOOPS_WEBHOOK_SUCCEEDED",
              processingStatus: "webhook_accepted",
              durationMs: elapsedMs(),
              response: response.data,
          });

          // If there's card JSON data in the response, create an adaptive card.
          if (response.data && response.data.cardJson) {
              const card = CardFactory.adaptiveCard(JSON.parse(response.data.cardJson));

              // If action type is not for fetching input, display card.
              if (response.data.actionType !== 'FetchInputAdaptiveCard') {
                  if (responseAsReply) {
                      await sendTeamsActivity(context, {
                          type: 'message',
                          id: activityId,
                          attachments: [card]
                      });
                  } else {
                      await updateTeamsActivity(context, {
                          type: 'message',
                          id: activityId,
                          attachments: [card]
                      });
                  }
              }
          }
          // If there's a message in the response, send it.
          else if (response.data && response.data.message) {
              if (responseAsReply) {
                  await sendTeamsActivity(context, {
                      type: 'message',
                      id: context.activity.replyToId,
                      text: response.data.message
                  });
              } else {
                  await updateTeamsActivity(context, {
                      type: 'message',
                      id: context.activity.replyToId,
                      text: response.data.message
                  });
              }
          }
      }
      // If backend response indicates an error, log it and inform the user.
      else {
          logger.error("Backend response was not ok", {
              event: "LOOPS_WEBHOOK_FAILED",
              processingStatus: "webhook_failed",
              durationMs: elapsedMs(),
              response: response.data,
              activityId: context.activity && context.activity.id,
          });
          const defaultCardJson = getDefaultAdaptiveCard(response.data.cardJson);
          const card = CardFactory.adaptiveCard(defaultCardJson);
          if (responseAsReply) {
              await sendTeamsActivity(context, {
                  type: 'message',
                  id: context.activity.replyToId,
                  attachments: [card],
                  channelData: {
                      notification: {
                          alert: true
                      },
                      microsoftTeams: {
                          entityID: context.activity.replyToId,
                          task: {
                              fetchTask: true,
                              type: "continue",
                              title: "Error",
                              height: 500,
                              width: 400
                          }
                      }
                  }
              });
          } else {
              await updateTeamsActivity(context, {
                  type: 'message',
                  id: context.activity.replyToId,
                  attachments: [card],
                  channelData: {
                      notification: {
                          alert: true
                      },
                      microsoftTeams: {
                          entityID: context.activity.replyToId,
                          task: {
                              fetchTask: true,
                              type: "continue",
                              title: "Error",
                              height: 500,
                              width: 400
                          }
                      }
                  }
              });
          }
      }
  } catch (error) {
      logger.error("Error processing the backend response", {
          event: "TEAMS_ACTIVITY_FAILED",
          durationMs: elapsedMs(),
          error,
      });
      await updateTeamsActivity(context, {
          type: 'message', // Activity type
          id: context.activity.replyToId, // ID of the activity to be updated
          text: 'An unexpected error occurred. Please try again later.' // New text to replace the previous message
      });
  }
}




/**
* The function uses a regular expression to identify and extract
* a messageId from a provided string. The messageId is assumed to be 
* a series of one or more numeric characters following the string "messageid=".
*/
function extractThreadId(inputString) {
  // Define a regular expression pattern to identify and extract the messageId.
  // The pattern looks for the string "messageid=", followed by one or more numeric characters.
  // The parentheses () create a capturing group, so that the numeric part can be extracted directly.
  const messageIdRegex = /messageid=([0-9]+)/;

  // Apply the regular expression to the input string.
  // If the pattern is found, match will be an array where match[1] is the first captured group (the numeric part).
  // If the pattern is not found, match will be null.
  const match = inputString.match(messageIdRegex);

  // Check if the pattern was found in the input string.
  if (match && match[1]) {
      // If a messageId was found, return it.
      return match[1];
  } else {
      // If a messageId was not found, return null.
      return null;
  }
}



// Export the LoopsCommandHandler class for use in your bot.
module.exports = {
  LoopsCommandHandler,
};