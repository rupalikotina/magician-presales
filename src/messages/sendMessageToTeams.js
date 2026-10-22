// src/messages/sendMessageToTeams.js
const { MessageFactory } = require("botbuilder");
const { AuthenticationConstants } = require("botframework-connector");
const { logger, elapsedMs } = require("../internal/logger");
const { sendTeamsActivity, updateTeamsActivity } = require("../internal/teamsResponse");

/**
 * Ensure we have a ConversationReference-like object the adapter can use.
 * Fills in bot.id from your AppId if it's missing.
 *
 * NOTE: For isUser=true (1:1/personal chat), we do NOT treat the raw AAD user id
 * as a conversation id. Bot Framework/Teams conversation ids for personal chats
 * (e.g. "a:1AbC2...@unq.gbl.spaces") are only obtained by calling
 * adapter.createConversation(), which needs `user`, `bot`, and `conversation.tenantId`.
 * We keep `user` on the ref so the caller can pass it through to createConversation.
 */
function normalizeConversationReference(input, appId, isUser) {
  const conv = input.conversation || {};
  const bot = input.bot || {};
  const user = input.user || {};
  const ref = {
    serviceUrl: input.serviceUrl,                  // e.g., https://smba.trafficmanager.net/amer/
    channelId: "msteams",
    bot: {
      id: bot.id || (appId ? `28:${appId}` : undefined),
      name: bot.name || "Bot",
    },
    user: user.id ? { id: user.id, name: user.name } : undefined,
    conversation: {
      // Only meaningful for channel/group sends; for isUser=true this is resolved
      // by adapter.createConversation() and must not be pre-filled with user.id.
      id: isUser ? undefined : conv.id,              // '19:...@thread.tacv2'
      tenantId: conv.tenantId || input.tenantId,  // Teams tenant GUID
      isGroup: !isUser,
      conversationType: conv.conversationType || (isUser ? "personal" : "channel"),
    },
  };
  return ref;
}

/**
 * Walks an arbitrary object/array tree and, wherever it finds a key literally named
 * `cardActivityId`, overwrites it with `activityId`. This is intentionally schema-agnostic -
 * it doesn't know or care whether that key lives inside `selectAction.data` on a button,
 * how deep it's nested, or how many buttons/actions carry it - any card, present or future,
 * that wants to know its own activity id just needs a field with this name somewhere in its
 * action data, and it'll get filled in.
 * @returns {boolean} true if at least one `cardActivityId` field was found and set.
 */
function embedSelfActivityId(node, activityId) {
  let found = false;
  const walk = (n) => {
    if (!n || typeof n !== "object") return;
    if (Array.isArray(n)) {
      n.forEach(walk);
      return;
    }
    if (Object.prototype.hasOwnProperty.call(n, "cardActivityId")) {
      n.cardActivityId = activityId;
      found = true;
    }
    Object.values(n).forEach(walk);
  };
  walk(node);
  return found;
}

/**
 * Send or update a message in a Teams channel using the BotFramework Adapter.
 * IMPORTANT: The adapter will mint a Bot Framework token using your bot's AppId/secret.
 *            Any client Authorization header is ignored at this layer.
 *
 * @param {CloudAdapter} adapter
 * @param {object} conversationReference - must contain serviceUrl + conversation.id (thread) + tenantId
 * @param {object} [cardContent]         - optional attachment (e.g., Adaptive Card)
 * @param {string} [messageText]         - optional text
 * @param {string} [activityId]          - when updateMode=true, the activity id to update
 * @param {boolean} [isUser]             - if true, uses the user-specific conversation id when present
 * @param {boolean} [updateMode]         - if true, updates an existing activity
 * @returns {Promise<object>}
 */
async function sendMessageToTeams(
  adapter,
  conversationReference,
  cardContent,
  messageText,
  activityId,
  isUser,
  updateMode
) {
  const appId = process.env.MicrosoftAppId || process.env.BOT_ID;

  // Build a safe reference
  const ref = normalizeConversationReference(conversationReference, appId, isUser);

  // Guardrails
  if (!ref.serviceUrl) throw new Error("conversationReference.serviceUrl is required.");
  if (!ref.conversation?.tenantId)
    throw new Error("conversationReference.conversation.tenantId (Teams tenant) is required.");
  if (!ref.bot?.id)
    throw new Error("Bot id missing. Ensure manifest botId and MicrosoftAppId match the installed bot.");
  if (isUser) {
    if (!ref.user?.id)
      throw new Error("conversationReference.user.id is required to message a user directly (isUser=true).");
  } else if (!ref.conversation?.id) {
    throw new Error("conversationReference.conversation.id (threadId) is required.");
  }

  let result = {};

  const runTurn = async (context) => {
    // Build activity
    const activity = cardContent
      ? { type: "message", attachments: [cardContent] }
      : MessageFactory.text(messageText ?? "");

    // Teams channel targeting (thread) is only relevant for channel/group sends.
    // For 1:1 sends the conversation id comes from createConversation(), not channelData.
    if (!isUser) {
      activity.channelData = {
        channel: { id: ref.conversation.id },
        tenant: { id: ref.conversation.tenantId },
      };
    }

    if (updateMode && activityId) {
      const updateResp = await updateTeamsActivity(context, { id: activityId, ...activity });
      result = { mode: "update", activityId, updateResp };
    } else {
      const sendResp = await sendTeamsActivity(context, activity); // returns { id: '<activityId>' }
      result = { mode: "send", activityId: sendResp?.id };

      // If the card carries a `cardActivityId` field anywhere in its action data (the
      // fallback teamsBot.js uses when Teams doesn't set replyToId on button taps - which
      // is the normal case for proactively-sent 1:1 cards), fill it in with the real id we
      // just got back and rewrite the message in place. This happens in the same turn, so
      // callers get a self-sufficient card from a single request - no manual two-step send.
      if (cardContent && sendResp?.id) {
        const embeddedCard = JSON.parse(JSON.stringify(cardContent));
        if (embedSelfActivityId(embeddedCard, sendResp.id)) {
          await updateTeamsActivity(context, {
            id: sendResp.id,
            type: "message",
            attachments: [embeddedCard],
          });
          result.cardActivityIdEmbedded = true;
        }
      }
    }
  };

  if (isUser) {
    // Personal (1:1) chat: the real conversation id must be minted by Bot Framework via
    // createConversation() using {bot, user, conversation.tenantId}. Passing the raw AAD
    // user id as conversation.id (the old behavior) is invalid and causes the send to fail.
    // CloudAdapter (UserAssignedMSI) requires createConversationAsync / continueConversationAsync.
    await adapter.createConversationAsync(
      appId,
      ref.channelId,
      ref.serviceUrl,
      AuthenticationConstants.ToChannelFromBotOAuthScope,
      {
        isGroup: false,
        bot: ref.bot,
        members: [ref.user],
        tenantId: ref.conversation.tenantId,
        channelData: { tenant: { id: ref.conversation.tenantId } },
      },
      runTurn
    );
  } else {
    await adapter.continueConversationAsync(appId, ref, runTurn);
  }

  const details = {
    ok: true,
    ...result,
    serviceUrl: ref.serviceUrl,
    threadId: ref.conversation.id,
    tenantId: ref.conversation.tenantId,
  };
  logger.info("Teams message sent", {
    event: "TEAMS_RESPONSE_SUCCEEDED",
    durationMs: elapsedMs(),
    isUser: Boolean(isUser),
    updateMode: Boolean(updateMode),
    messageText,
    cardContent,
    details,
  });
  return details;
}

module.exports = { sendMessageToTeams };
