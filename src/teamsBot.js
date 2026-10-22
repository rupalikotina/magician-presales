// Required libraries and modules
const {
    TeamsActivityHandler,
    MessageFactory,
    InputHints,
    CardFactory,
    TeamsInfo
} = require('botbuilder');
const { postLoops } = require('./internal/loopsClient');
const { logger, elapsedMs } = require('./internal/logger');
const { sendTeamsActivity, updateTeamsActivity, logConnectorError } = require('./internal/teamsResponse');
const { noteActivity } = require('./internal/duplicateDetector');

// Constants
const BACKEND_URL = 'https://api.presalestest.theloops.ai/conversation/v1/webhooks/msteams/events';

// IN-MEMORY vote tracker, keyed by the card's activity id -> Map(aadObjectId -> {name, verb}).
// This is concept/demo-only: it lives in this bot process's memory, so it resets on restart and
// won't be shared across multiple bot instances if this ever scales out horizontally. It also
// isn't visible to core/integrations for reporting. Once TLDV-3367's backend processing of
// "MessageFeedback" events exists, this aggregation should move there (bot would fetch the
// current vote list from the backend before rebuilding the card, instead of keeping it here).
const feedbackVotesByActivity = new Map();

/**
 * Records one user's vote on a given card activity, then returns the current aggregate
 * (who liked it, who disliked it) plus whether this tap was an undo.
 *
 * Three cases per tap:
 *  - No prior vote from this user -> set it (vote added).
 *  - Prior vote from a DIFFERENT verb -> overwrite it (switched, not double-counted).
 *  - Prior vote from the SAME verb (tapping the button you already picked again) -> remove
 *    it entirely (undo) rather than just re-confirming the same value. Without this, there
 *    would be no way to back out of a vote back to "no opinion" - only to switch sides.
 */
function recordVoteAndSummarize(activityId, from, verb) {
    if (!activityId) {
        return { likers: [], dislikers: [], undone: false };
    }
    const voterKey = (from && (from.aadObjectId || from.id)) || undefined;
    let undone = false;
    if (voterKey) {
        if (!feedbackVotesByActivity.has(activityId)) {
            feedbackVotesByActivity.set(activityId, new Map());
        }
        const votesForActivity = feedbackVotesByActivity.get(activityId);
        const existing = votesForActivity.get(voterKey);
        if (existing && existing.verb === verb) {
            votesForActivity.delete(voterKey);
            undone = true;
        } else {
            votesForActivity.set(voterKey, { name: (from && from.name) || 'Someone', verb });
        }
    }

    const votes = feedbackVotesByActivity.get(activityId);
    const likers = [];
    const dislikers = [];
    if (votes) {
        for (const vote of votes.values()) {
            (vote.verb === 'like' ? likers : dislikers).push(vote.name);
        }
    }
    return { likers, dislikers, undone };
}

// TeamsBot class that handles interactions within MS Teams
class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();


        // Event handler for when a new message is received.
        this.onMessage(async (context, next) => {
            try {
                try {
                    noteActivity(context && context.activity, { suppress: false });
                } catch (error) {
                    logger.error("Duplicate observation failed", {
                        event: "TEAMS_ACTIVITY_FAILED",
                        error,
                    });
                }

                // Guard: let TeamsFx ConversationBot handle 'loops'/'ask' commands; skip generic handlers to avoid duplicates.
                try {
                    const raw = (context && context.activity && context.activity.text) ? String(context.activity.text) : '';
                    const text = raw.trim();
                    if (/^(loops|ask)\b/i.test(text)) {
                        logger.info("Delegating message to command handler", {
                            event: "TEAMS_ACTIVITY_CLASSIFIED",
                            actionType: "Command",
                            text,
                            activityId: context.activity && context.activity.id,
                        });
                        await next();
                        return; // do not continue to generic notifier paths
                    }
                } catch (e) {
                    logger.error("command guard failed", {
                        event: "TEAMS_ACTIVITY_FAILED",
                        error: e,
                    });
                }

                if (context.activity) {
                    // Extract Adaptive Card submit data
                    const formData = context.activity.value;

                    if (context.activity && context.activity.text) {
                        const text = context.activity.text.trim().toLowerCase();
                        logger.info("Received message", {
                            event: "TEAMS_ACTIVITY_RECEIVED",
                            activityType: context.activity.type,
                            text,
                            activityId: context.activity.id,
                        });

                        // Check if the message is in a personal chat
                        if (context.activity.conversation.conversationType === "personal") {
                            logger.info("Classified personal message", {
                                event: "TEAMS_ACTIVITY_CLASSIFIED",
                                actionType: "UserInteractionWithBot",
                                text,
                                activityId: context.activity.id,
                                conversation: context.activity.conversation,
                            });
                            // Notify the backend that a new message was received in a direct chat with the bot.
                            await notifyBackendForNewMessage('UserInteractionWithBot', context);
                        } else {
                            // Check if the message mentions the bot.
                            const isBotMentioned = context.activity.entities.some(entity => entity.type === 'mention' && entity.mentioned.id === context.activity.recipient.id);

                            // Determine the appropriate action based on the message content.
                            if (isBotMentioned) {
                                const words = text.split(' ');
                                const loopsIndex = words.indexOf("loops");

                                if (loopsIndex === -1) {
                                    // Notify the backend that a new message was received.
                                    const actionType = 'LoopsBotMentioned';
                                    logger.info("Classified bot mention", {
                                        event: "TEAMS_ACTIVITY_CLASSIFIED",
                                        actionType,
                                        text,
                                        activityId: context.activity.id,
                                    });
                                    await notifyBackendForNewMessage(actionType, context);
                                } else {
                                    logger.info("Bot mention left to command handler", {
                                        event: "TEAMS_ACTIVITY_CLASSIFIED",
                                        actionType: "Command",
                                        text,
                                        activityId: context.activity.id,
                                    });
                                }
                            } else {
                                // If the bot is not mentioned, still notify the backend about the new message.
                                const conversationId = context.activity.conversation.id;
                                const activityId = context.activity.id;

                                const messageIdMatch = conversationId.match(/messageid=(\d+)/);
                                const actionType = ((messageIdMatch && messageIdMatch[1]) && messageIdMatch[1] !== activityId) ? 'NewThreadedMessage' : 'NewMessage';
                                logger.info("Classified channel message", {
                                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                                    actionType,
                                    text,
                                    activityId,
                                    conversationId,
                                });
                                await notifyBackendForNewMessage(actionType, context);
                            }
                        }
                        await next();
                    } else if (formData) {
                        try {
                            if (formData.verb === 'like' || formData.verb === 'dislike') {
                                logger.info("Classified card feedback", {
                                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                                    actionType: "MessageFeedback",
                                    verb: formData.verb,
                                    formData,
                                    activityId: context.activity.id,
                                });
                                // Like/dislike tap arriving as an Action.Submit (activity.value).
                                await handleCardFeedback(context, formData);
                            } else if (formData.dropdownChanged) {
                                const actionType = 'UpdateAdaptiveCardOnDropDown';
                                logger.info("Activity classified", {
                                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                                    actionType,
                                    formData,
                                    activityId: context.activity.id,
                                });
                                const dropdownValue = cardInput.dropdown;

                                // Create the payload
                                const payload = {
                                    ...await constructPayloadForBackend(actionType, formData, context),
                                    selectedValue: dropdownValue
                                };

                                // Update the card based on the dropdownValue
                                const updatedCard = await generateUpdatedCard(payload);
                                await updateTeamsActivity(context, {
                                    attachments: [updatedCard]
                                });
                            } else {
                                const actionType = 'ProcessAdaptiveCardForm';
                                logger.info("Activity classified", {
                                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                                    actionType,
                                    formData,
                                    activityId: context.activity.id,
                                });
                                // Create the payload
                                const payload = {
                                    ...await constructPayloadForBackend(actionType, formData, context),
                                    from: context.activity.from,
                                    to: context.activity.recipient,
                                    threadId: extractThreadId(context.activity.conversation.id),
                                    eventTimestamp: context.activity.timestamp,
                                    eventId: context.activity.id,
                                    hiddenData: context.activity.value.hiddenData
                                };

                                // Default 'isButton' and 'responseAsReply' to false if not present
                                // Extract buttonData
                                const buttonData = formData.buttonData || {};

                                // Default 'isButton' and 'responseAsReply' to false if not present
                                const isButton = buttonData.isButton || false;
                                const responseAsReply = buttonData.responseAsReply || false;

                                const payloadJsonString = JSON.stringify(payload);
                                const response = await postLoops(BACKEND_URL, payloadJsonString);
                                await processBackendResponse(response, context, responseAsReply);
                            }
                        } catch (error) {
                            logger.error("Error in onMessage", {
                                event: "TEAMS_ACTIVITY_FAILED",
                                error,
                            });
                            await updateTeamsActivity(context, error.message);
                        }
                    }
                }

            } catch (error) {
                logger.error("Error in onMessage", {
                    event: "TEAMS_ACTIVITY_FAILED",
                    error,
                });
                await sendTeamsActivity(context, error.message);
            }
        });



        // Event handler for when a new channel is created.
        this.onConversationUpdate(async (context, next) => {
            try {
                try {
                    noteActivity(context && context.activity, { suppress: false });
                } catch (error) {
                    logger.error("Duplicate observation failed", {
                        event: "TEAMS_ACTIVITY_FAILED",
                        error,
                    });
                }

                logger.info("Conversation update received", {
                    event: "TEAMS_ACTIVITY_RECEIVED",
                    activityType: context.activity.type,
                    eventType: context.activity.channelData && context.activity.channelData.eventType,
                    activityId: context.activity.id,
                    membersAdded: context.activity.membersAdded,
                    membersRemoved: context.activity.membersRemoved,
                    channelData: context.activity.channelData,
                });
                if (context.activity.channelData.eventType == "teamMemberAdded") {
                    await handleTeamMembersAdded(context);
                } else if (context.activity.channelData.eventType == "teamMemberRemoved") {
                    await handleTeamMembersRemoved(context);
                }

                await next();
            } catch (error) {
                logger.error("Error in onConversationUpdate", {
                    event: "TEAMS_ACTIVITY_FAILED",
                    error,
                });
            }
        });

    }

    /**
     * IMPORTANT: Action.Execute buttons (Universal Actions, e.g. our like/dislike
     * selectActions) are delivered as an Invoke activity named 'adaptiveCard/action'.
     * The installed botbuilder version (4.23.3) has NO built-in case for that name in
     * TeamsActivityHandler.onInvokeActivity()'s switch statement, so it falls through to
     * the base ActivityHandler, which has nothing registered for it either. Teams then
     * shows "Unable to reach app. Please try again." because it never gets a valid
     * InvokeResponse back.
     *
     * Previously this bot tried to handle this via `this.onInvoke = async (context, next) => {...}`
     * in the constructor - that was a no-op: TeamsActivityHandler has no `onInvoke`
     * registration hook (unlike `onMessage`), so that function was never called by the
     * framework's dispatch logic. Overriding this actual method is required.
     *
     * You CAN call any API directly from in here (this is exactly where a like/dislike
     * tap should trigger your backend call/event) - just make sure to `return` the
     * InvokeResponse instead of using context.sendActivity(...).
     */
    async onInvokeActivity(context) {
        try {
            noteActivity(context && context.activity, { suppress: false });
        } catch (error) {
            logger.error("Duplicate observation failed", {
                event: "TEAMS_ACTIVITY_FAILED",
                error,
            });
        }
        if (context.activity.name === 'adaptiveCard/action') {
            try {
                const invokeValue = context.activity.value; // { action: { type, id, verb, data } }
                const verb = invokeValue && invokeValue.action ? invokeValue.action.verb : undefined;
                const actionData = invokeValue && invokeValue.action ? invokeValue.action.data : undefined;

                logger.info("Adaptive card invoke received", {
                    event: "TEAMS_ACTIVITY_RECEIVED",
                    activityType: context.activity.type,
                    verb,
                    actionData,
                    activityId: context.activity.id,
                    from: context.activity.from,
                    conversation: context.activity.conversation,
                });
                logger.info("Activity classified", {
                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                    actionType: verb || context.activity.name,
                    verb,
                    actionData,
                    activityId: context.activity.id,
                });
                switch (verb) {
                    case 'like':
                    case 'dislike': {
                        // Call the backend directly, synchronously, as part of handling the tap.
                        const feedbackBody = {
                            action: 'MessageFeedback',
                            feedback: verb,
                            data: actionData,
                            from: context.activity.from,
                            to: context.activity.recipient,
                            eventId: context.activity.id,
                            eventTimestamp: context.activity.timestamp
                        };
                        await postLoops(BACKEND_URL, feedbackBody, { headers: { 'Content-Type': 'application/json' } });

                        return {
                            status: 200,
                            body: {
                                statusCode: 200,
                                type: 'application/vnd.microsoft.card.adaptive',
                                value: {
                                    type: 'AdaptiveCard',
                                    version: '1.4',
                                    $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
                                    body: [{
                                        type: 'TextBlock',
                                        wrap: true,
                                        text: verb === 'like' ? '👍 Thanks for the feedback!' : '👎 Thanks — we\'ll do better.'
                                    }]
                                }
                            }
                        };
                    }
                    case 'submit':
                        await this.handleCardSubmitAction(context, actionData);
                        return { status: 200, body: { statusCode: 200 } };
                    default:
                        logger.error("Unknown adaptiveCard/action verb", {
                            event: "TEAMS_ACTIVITY_FAILED",
                            verb,
                            activityId: context.activity.id,
                        });
                        return { status: 200, body: { statusCode: 400, type: 'application/vnd.microsoft.error', value: { code: 'UnknownVerb', message: `Unknown action verb: ${verb}` } } };
                }
            } catch (error) {
                logger.error("Error handling adaptiveCard/action invoke", {
                    event: "TEAMS_ACTIVITY_FAILED",
                    error,
                });
                return { status: 500, body: { statusCode: 500, type: 'application/vnd.microsoft.error', value: { code: 'InternalError', message: error.message } } };
            }
        }

        // Let TeamsActivityHandler continue to handle every other invoke name
        // (task/fetch, task/submit, config/fetch, composeExtension/*, etc.) as normal.
        return super.onInvokeActivity(context);
    }

    // Function to handle the card submit action
    async handleCardSubmitAction(context, actionData) {
        // Process actionData here (e.g., save to database, call an API, etc.)
        // For now, let's log the data
        logger.info("Received submit action data", {
            event: "TEAMS_CARD_SUBMITTED",
            actionData,
            activityId: context.activity.id,
        });

        // Optionally send a response message to the user
        await sendTeamsActivity(context, 'Your form data was received and processed.');
    }

}



/**
 * Notifies the backend with a new message by sending a POST request.
 */
async function notifyBackendForNewMessage(actionType, context) {

    // Check if the bot is mentioned
    const isBotMentioned = context.activity.entities && context.activity.entities.some(entity => entity.type === 'mention' && entity.mentioned.id === context.activity.recipient.id);

    // Construct the payload to be sent to the backend.
    const payload = {
        ...await constructPayloadForBackend(actionType, null, context),
        messageText: context.activity.text,
        eventId: context.activity.id,
        threadId: extractThreadId(context.activity.conversation.id),
        eventTimestamp: context.activity.timestamp,
        from: context.activity.from,
        to: context.activity.recipient,
        isBotMentioned: isBotMentioned
    };
    const payloadJsonString = JSON.stringify(payload);

    try {
        // Sending a POST request to the backend with the constructed payload.
        // postLoops logs start, acceptance, and failure without changing this error.
        await postLoops(BACKEND_URL, payloadJsonString, {
            headers: {
                'Content-Type': 'application/json'
            }
        });
    } catch (error) {
        // Swallow so a webhook failure does not fail the Teams turn. postLoops already logged it.
        logger.error(`Error notifying backend for ${actionType}`, {
            event: "TEAMS_ACTIVITY_FAILED",
            actionType,
            durationMs: elapsedMs(),
            error,
        });
    }
}



/**
 * Handles a like/dislike tap on a card's feedback ColumnSet (Action.Submit -> normal
 * message activity with activity.value). Makes the choice "stick" two ways:
 *
 *  1. Visually: rewrites the original card via context.updateActivity() so the
 *     selection is baked into the actual stored message - it survives reloads,
 *     scrolling away and back, or reopening the chat, because the message itself
 *     was edited (not just re-rendered client-side). Note this replaces the card
 *     for *everyone* who can see that message (channel/group chat), since Teams
 *     messages don't have a per-viewer rendering - there's no way to show user A
 *     "you voted 👍" while user B still sees the original buttons.
 *  2. Durably: posts the vote to the backend (BACKEND_URL) so it's recorded outside
 *     Teams entirely - queryable, auditable, and immune to the card ever being
 *     edited/deleted. This is the "save context" piece if you need to know who
 *     voted what later, prevent double-votes, etc.
 */
async function handleCardFeedback(context, formData) {
    const verb = formData.verb; // 'like' | 'dislike'

    // `replyToId` on an Action.Submit activity is set by Teams to the id of the message
    // that contained the card - that's the ONLY reliable id to key the vote/update the card
    // by. Do not fall back to formData.messageId: that value comes from whatever the card's
    // own JSON had baked into its selectAction.data, which in our test cards is still the
    // literal unsubstituted string "${messageId}" - not a real activity id.
    logger.info("handleCardFeedback", {
        event: "TEAMS_ACTIVITY_RECEIVED",
        verb,
        activityId: context.activity.id,
        replyToId: context.activity.replyToId,
        conversationId: context.activity.conversation && context.activity.conversation.id,
    });

    // Prefer replyToId (the Bot Framework-correct answer to "which message was this card
    // on"), but fall back to formData.cardActivityId - a field the SENDER is responsible
    // for baking into the card's own button data, using the real activity id that
    // sendMessageToTeams.js already returns from the initial send. This matters for cards
    // sent proactively (via adapter.createConversation(), not as a reply within a live
    // turn) - replyToId has not been reliably observed to be set by Teams for that path.
    // See the two-step send pattern: send the card once, take the returned activityId from
    // the response, then re-send with updateMode:true and cardActivityId set to that id.
    const originalActivityId = context.activity.replyToId || formData.cardActivityId;
    if (!originalActivityId) {
        logger.error("handleCardFeedback: no replyToId AND no formData.cardActivityId - cannot record or update the original card", {
            event: "TEAMS_RESPONSE_FAILED",
            verb,
            activityId: context.activity.id,
            replyToId: context.activity.replyToId,
            conversationId: context.activity.conversation && context.activity.conversation.id,
            formData,
        });
        return;
    }

    // Record this user's vote FIRST - overwriting a prior different vote (switch), or
    // removing it entirely if they tapped the same button again (undo) - then get back
    // who's currently on each side. Everything downstream (backend event, rebuilt card)
    // is derived from this single source of truth, so it can't disagree with itself.
    const voteSummary = recordVoteAndSummarize(originalActivityId, context.activity.from, verb);

    try {
        // Route through the SAME envelope every other Teams event uses (channelId,
        // teamsTenantId, teamId, teamName at the top level, business payload nested under
        // `data`). This was previously built ad hoc here - flat, and missing channelId/
        // teamId/teamsTenantId entirely - which meant core's webhook handler had no way to
        // resolve which ServiceIntegration this event belonged to (it needs a teamId, or a
        // channelId it can map to one) and threw "No ServiceIntegration found for teamId"
        // even for a well-formed vote. formData (messageId, entityId, commentId, ticketId,
        // whatever gets added later) still travels opaquely inside `data` - teamsBot.js
        // still doesn't need to know or enumerate the card's schema.
        const backendPayload = await constructPayloadForBackend('MessageFeedback', formData, context);
        const feedbackPayload = {
            ...backendPayload,
            feedback: verb,
            undone: voteSummary.undone, // true = user removed their own vote, not added/changed it
            from: context.activity.from,
            to: context.activity.recipient,
            eventId: context.activity.id,
            eventTimestamp: context.activity.timestamp
        };
        await postLoops(BACKEND_URL, feedbackPayload, { headers: { 'Content-Type': 'application/json' } });
    } catch (error) {
        logger.error("Failed to record message feedback in backend", {
            event: "TEAMS_ACTIVITY_FAILED",
            durationMs: elapsedMs(),
            error,
        });
    }

    // Rebuild the SAME card, still fully interactive, re-highlighted purely from the
    // aggregate above - this is what lets the user tap the other emoji to switch, or
    // tap their own vote again to undo it, instead of the buttons disappearing after
    // the first tap. Whatever extra keys formData had (commentId, ticketId, ...) get
    // carried forward untouched so the next tap still has them.
    const resultCard = buildFeedbackCard(formData, voteSummary);

    try {
        await updateTeamsActivity(context, {
            type: 'message',
            id: originalActivityId,
            attachments: [CardFactory.adaptiveCard(resultCard)]
        });
    } catch (_error) {
        // updateTeamsActivity already logged TEAMS_RESPONSE_FAILED or TEAMS_AUTH_FAILED
        // with status and body. Swallow so a card rewrite failure does not undo the vote.
    }
}



/**
 * Builds the like/dislike feedback ROW and splices it onto whatever the original card's
 * body already was. This function has zero opinion on what the rest of the card looks like
 * (title wording, layout, extra fields) - that content comes from the sender and is carried
 * through untouched, because the payload isn't fixed and will keep changing. Hardcoding a
 * reconstruction of "the top of the card" here would mean every rewrite silently throws away
 * whatever the sender actually put there and replaces it with a fixed stand-in.
 *
 * Both emoji stay live/tappable no matter the current state - tapping either re-runs
 * handleCardFeedback, which calls recordVoteAndSummarize() BEFORE this function, so
 * `voteSummary` always already reflects the post-tap truth (added, switched, or undone).
 * That's deliberate: this function never infers state from "which button was just tapped" -
 * it only ever renders whatever the aggregate currently says, so an undo (tapping your own
 * vote again to remove it) renders correctly instead of getting stuck highlighted.
 *
 * @param {object} data - the ORIGINAL submitted payload (or whatever the card was first
 *   sent with), schema-agnostic. Every key on it (messageId, entityId, commentId, ticketId,
 *   `cardActivityId`, or anything added later) is carried forward as-is into both buttons'
 *   selectAction.data, so this function never needs to be touched when new fields get added
 *   upstream. `data.cardBody`, if present, is an array of Adaptive Card elements - everything
 *   the sender wants ABOVE the feedback row (title, ticket details, whatever) - and is
 *   reproduced verbatim. If it's missing, the card is just the feedback row on its own.
 * @param {{likers: string[], dislikers: string[]}} [voteSummary] - everyone who currently
 *   has a vote on this card. Since the card is shared by every viewer (channel/group chat),
 *   "selected" here means "at least one person currently votes this way", not "the current
 *   viewer voted this way" - there's no such thing as a per-viewer render.
 */
function buildFeedbackCard(data, voteSummary) {
    const { cardBody } = data || {};
    const likers = (voteSummary && voteSummary.likers) || [];
    const dislikers = (voteSummary && voteSummary.dislikers) || [];
    const likeSelected = likers.length > 0;
    const dislikeSelected = dislikers.length > 0;

    // Re-embed everything the card came in with, just overriding `verb`/`feedback` per button.
    const likeData = { ...data, verb: 'like', feedback: 'like' };
    const dislikeData = { ...data, verb: 'dislike', feedback: 'dislike' };

    const feedbackRow = {
        type: 'ColumnSet',
        id: 'feedbackEmojis',
        spacing: 'Medium',
        columns: [
            {
                type: 'Column',
                width: 'auto',
                items: [
                    {
                        // Native Adaptive Card Icon (Fluent icon set, schema 1.5+) - this is
                        // what actually gives an outline vs. filled glyph, unlike a plain
                        // emoji TextBlock which has no stroke/outline property at all.
                        // style: "Regular" = outline (not selected), "Filled" = solid (selected).
                        type: 'Icon',
                        name: 'ThumbLike',
                        size: 'xSmall',
                        style: likeSelected ? 'Filled' : 'Regular',
                        // Adaptive Cards has no literal "Yellow" in its color enum (Default,
                        // Dark, Light, Accent, Good, Warning, Attention) - "Warning" is the
                        // closest built-in and is what Teams actually renders as yellow/amber.
                        color: likeSelected ? 'Warning' : 'Default'
                    }
                ],
                selectAction: { type: 'Action.Submit', id: 'likeAction', data: likeData }
            },
            {
                type: 'Column',
                width: 'auto',
                items: [
                    {
                        type: 'Icon',
                        name: 'ThumbDislike',
                        size: 'xSmall',
                        style: dislikeSelected ? 'Filled' : 'Regular',
                        color: dislikeSelected ? 'Warning' : 'Default'
                    }
                ],
                selectAction: { type: 'Action.Submit', id: 'dislikeAction', data: dislikeData }
            }
        ]
    };

    // No visible name list and no hint text - Adaptive Cards have no hover interaction at
    // all (tap/click only, and hover isn't even a concept on mobile), so a tooltip-style
    // reactor list isn't achievable here. `likers`/`dislikers` are still tracked (via
    // recordVoteAndSummarize) purely to decide each icon's fill state below - just not
    // rendered as text.

    // Whatever the sender put above the feedback row, reproduced as-is. No fallback title/
    // text reconstruction here - if the sender didn't include cardBody, the card is just
    // the feedback row.
    const body = [
        ...(Array.isArray(cardBody) ? cardBody : []),
        feedbackRow
    ];

    return {
        type: 'AdaptiveCard',
        version: '1.5', // bumped from 1.4 for showBorder/roundedCorners on the selected column
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        body
    };
}



/**
 * Constructs a payload to be sent to the backend containing
 */
async function constructPayloadForBackend(actionType, data, context) {
    // Extract channel details from the context.
    const channelId = context.activity.channelData && context.activity.channelData.channel ? context.activity.channelData.channel.id : 'unknown';
    let teamId = context.activity.channelData.team ? context.activity.channelData.team.id : undefined;
    let teamName;

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
 * Event handler for when a user joins the team.
 */
async function handleTeamMembersAdded(context) {
    try {
        const membersAdded = context.activity.membersAdded;

        // Iterate through all members that have been added.
        for (let member of membersAdded) {

            // Check if the bot itself is the member that was added.
            if (member.id === context.activity.recipient.id) {

                // Extract relevant details about the bot and the team.
                const botName = context.activity.recipient.name;
                let teamName = context.activity.channelData.team ? context.activity.channelData.team.name : "Unknown";

                // Define an action to fetch the welcome message for the bot.
                let actionType = 'FetchWelcomeMessage';
                logger.info("Activity classified", {
                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                    actionType,
                    activityType: context.activity.type,
                    memberId: member.id,
                    activityId: context.activity.id,
                });
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.id : undefined;

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
                const payload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const payloadJsonString = JSON.stringify(payload);

                // Attempt to fetch the welcome message for the bot from the backend.
                try {
                    const welcomeResponse = await postLoops(BACKEND_URL, payloadJsonString, {
                        headers: {
                            'Content-Type': 'application/json'
                        }
                    });

                    // Extract the welcome message or use a default message if none is provided.
                    let welcomeMessage = welcomeResponse.data && welcomeResponse.data.message ?
                        welcomeResponse.data.message :
                        `<at>${botName}</at> was added to ${teamName} by an administrator`;

                    // Craft a message with mentions and send it.
                    const message = MessageFactory.text(welcomeMessage, welcomeMessage, InputHints.ExpectingInput);
                    message.entities = [{
                        mentioned: context.activity.recipient,
                        text: `<at>${botName}</at>`,
                        type: 'mention'
                    }];
                    await sendTeamsActivity(context, message);

                    // Notify the backend about the addition of the bot to a team or channel.
                    actionType = 'NotifyBotAddition';
                    logger.info("Activity classified", {
                        event: "TEAMS_ACTIVITY_CLASSIFIED",
                        actionType,
                        activityId: context.activity.id,
                    });
                    let teamId = context.activity.channelData.team ? context.activity.channelData.team.aadGroupId : undefined;

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
                    const notifyPayload = {
                        action: actionType,
                        teamsTenantId: context.activity.channelData.tenant.id,
                        teamId: teamId,
                        teamName: teamName,
                        channelId: context.activity.channelData.channel.id,
                        from: context.activity.from,
                        to: context.activity.recipient,
                        rawEvent: context
                    };

                    const notifyPayloadJsonString = JSON.stringify(notifyPayload);
                    await postLoops(BACKEND_URL, notifyPayloadJsonString, {
                        headers: {
                            'Content-Type': 'application/json'
                        }
                    });

                } catch (error) {
                    // If there's an error fetching the welcome message, log it and send a default message.
                    // postLoops already recorded LOOPS_WEBHOOK_FAILED for the HTTP call.
                    logger.error("Error fetching welcome message from backend", {
                        event: "TEAMS_ACTIVITY_FAILED",
                        durationMs: elapsedMs(),
                        error,
                    });
                    if (defaultMessage) {
                        await sendTeamsActivity(context, defaultMessage);
                    }
                }
            } else {
                // For members other than the bot.
                const actionType = 'MemberJoinedChannel';
                logger.info("Activity classified", {
                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                    actionType,
                    userId: member.aadObjectId,
                    activityId: context.activity.id,
                });
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.aadGroupId : undefined;

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
                const payload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    userId: member.aadObjectId,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const payloadJsonString = JSON.stringify(payload);

                // Notify the backend about the addition of a member to a team or channel.
                await postLoops(BACKEND_URL, payloadJsonString, {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                });
            }
        }
    } catch (error) {
        logger.error("Error in onMembersAdded", { event: "TEAMS_ACTIVITY_FAILED", error });
    }
}



/**
 * Event handler for when a user is removed from the team.
 */
async function handleTeamMembersRemoved(context) {
    try {
        const membersRemoved = context.activity.membersRemoved;

        // Iterate through all members that have been removed.
        for (let member of membersRemoved) {

            // Check if the bot itself is the member that was removed.
            if (member.id === context.activity.recipient.id) {
                const botName = context.activity.recipient.name;
                const teamName = context.activity.channelData.team ? context.activity.channelData.team.name : "a channel";
                const removerName = context.activity.from.name;

                // Craft a message indicating the bot was removed and by whom, then send it.
                const removalMessage = `<at>${botName}</at> was removed from ${teamName} by <at>${removerName}</at>.`;
                await sendTeamsActivity(context, removalMessage);

                // Notify the backend about the removal of the bot to a team or channel.
                actionType = 'NotifyBotRemoval';
                logger.info("Activity classified", {
                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                    actionType,
                    activityId: context.activity.id,
                });
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.aadGroupId : undefined;

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
                const notifyPayload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const notifyPayloadJsonString = JSON.stringify(notifyPayload);
                await postLoops(BACKEND_URL, notifyPayloadJsonString, {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                });
            } else {
                // For members other than the bot.
                const actionType = 'MemberLeftChannel';
                logger.info("Activity classified", {
                    event: "TEAMS_ACTIVITY_CLASSIFIED",
                    actionType,
                    userId: member.aadObjectId,
                    activityId: context.activity.id,
                });
                let teamId = context.activity.channelData.team ? context.activity.channelData.team.id : undefined;

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
                const payload = {
                    action: actionType,
                    teamsTenantId: context.activity.channelData.tenant.id,
                    teamId: teamId,
                    teamName: teamName,
                    userId: member.aadObjectId,
                    from: context.activity.from,
                    to: context.activity.recipient,
                    rawEvent: context
                };

                const payloadJsonString = JSON.stringify(payload);

                // Notify the backend about the removal of a member from a team or channel.
                await postLoops(BACKEND_URL, payloadJsonString, {
                    headers: {
                        'Content-Type': 'application/json'
                    }
                });
            }
        }
    } catch (error) {
        logger.error("Error in onMembersRemoved", { event: "TEAMS_ACTIVITY_FAILED", error });
    }
}



/**
 * Event handler for when a channel is created for the team.
 */
async function handleTeamsChannelCreated(context) {
    try {
        const channelId = context.activity.channelData && context.activity.channelData.channel ? context.activity.channelData.channel.id : 'unknown';
        const channelName =  context.activity.channelData && context.activity.channelData.channel ? context.activity.channelData.channel.name : 'unknown';
        let teamId = context.activity.channelData && context.activity.channelData.team ? context.activity.channelData.team.id : undefined;

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

        // Create the payload to send to the backend.
        const actionType = 'NotifyChannelCreation';
        logger.info("Activity classified", {
            event: "TEAMS_ACTIVITY_CLASSIFIED",
            actionType,
            channelId,
            activityId: context.activity.id,
        });
        const payload = {
            action: actionType,
            teamsTenantId: context.activity.channelData.tenant.id,
            teamId: teamId, // Here you might want to use teamId directly, depending on your needs
            channelId: channelId,
            channelName: channelName,
            teamName: teamName,
            from: context.activity.from,
            to: context.activity.recipient,
            rawEvent: context
        };

        const payloadJsonString = JSON.stringify(payload);

        // Send the payload to the backend.
        await postLoops(BACKEND_URL, payloadJsonString, {
            headers: {
                'Content-Type': 'application/json'
            }
        });
    } catch (error) {
        logger.error("Error in onTeamsChannelCreated", { event: "TEAMS_ACTIVITY_FAILED", error });
    }
}




/**
 * Process the response from the backend and send the appropriate message/card
 * to the user based on the response content.
 */
async function processBackendResponse(response, context, responseAsReply) {
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
                            id: context.activity.replyToId,
                            attachments: [card]
                        });
                    } else {
                        await updateTeamsActivity(context, {
                            type: 'message',
                            id: context.activity.replyToId,
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
                                width: 400,
                                value: {
                                    processingActivityId: context.activity.replyToId // Storing the processingActivity.id in the value
                                }
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
                                width: 400,
                                value: {
                                    processingActivityId: context.activity.replyToId // Storing the processingActivity.id in the value
                                }
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
 * Generate an updated adaptive card by sending the given payload to the backend.
 * If the backend fails to provide an updated card, generate a default error card.
 */
async function generateUpdatedCard(payload) {
    // Send the payload to your backend and wait for a response.
    const payloadJsonString = JSON.stringify(payload);
    const response = await postLoops(YOUR_BACKEND_URL, payloadJsonString);
    logger.info("Backend response", {
        event: "LOOPS_WEBHOOK_SUCCEEDED",
        processingStatus: "webhook_accepted",
        durationMs: elapsedMs(),
        actionType: payload && payload.action,
        status: response.status,
        response: response.data,
    });

    // If the backend response is successful and contains card data, return the card.
    if (response.data.status === "ok" && response.data.cardJson) {
        return CardFactory.adaptiveCard(response.data.cardJson);
    } else {
        // If the backend response indicates an error or doesn't provide card data,
        // return a default error card.
        return CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            body: [{
                "type": "TextBlock",
                "text": "An error occurred while updating the card."
            }]
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



/**
 * Generates and returns the default Adaptive Card JSON structure.
 * This card is meant to be used as a fallback when the system is unable
 * to retrieve or render the primary form/data.
 */
function getDefaultAdaptiveCard(errorMessage) {
    return {
        // Specifies the type of the card.
        type: "AdaptiveCard",

        // Defines the version of the Adaptive Card schema that this card requires.
        version: "1.3",

        // Contains the primary card elements that compose the content of the card.
        body: [{
            // Text block element to show a bold title.
            type: "TextBlock",
            size: "Medium",
            weight: "Bolder",
            text: "Error"
        },
            {
                // Text block element to display the error message.
                type: "TextBlock",
                text: errorMessage,
                wrap: true // Allow the text to wrap and extend beyond one line.
            }
        ],

        // Specifies the schema this card uses. It's a standard practice to include this.
        $schema: "http://adaptivecards.io/schemas/adaptive-card.json"
    };
}



module.exports.TeamsBot = TeamsBot;