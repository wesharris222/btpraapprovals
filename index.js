const path = require('path');
const express = require('express');
const { BotFrameworkAdapter } = require('botbuilder');
const BTApprovalBot = require('./bot');

// Create adapter
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Create bot instance
const bot = new BTApprovalBot();

// Error handler
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    console.error('Error details:', {
        message: error.message,
        stack: error.stack,
        activity: context.activity
    });
    await context.sendActivity('The bot encountered an error or bug.');
};

// Create HTTP server
const app = express();
app.use(express.json());

// Listen for incoming activities
app.post('/api/messages', (req, res) => {
    console.log('Received message activity:', JSON.stringify(req.body, null, 2));
    adapter.processActivity(req, res, async (context) => {
        await bot.run(context);
    });
});

// Webhook endpoint for Logic App
app.post('/api/webhook', async (req, res) => {
    try {
        console.log('Webhook request received:', JSON.stringify(req.body, null, 2));

        const conversationReferences = await bot.getAllConversationReferences();
        if (!conversationReferences || conversationReferences.length === 0) {
            throw new Error('No conversation references found');
        }

        // Send message to all stored conversations
        for (const reference of conversationReferences) {
            try {
                console.log('Sending to conversation:', reference.conversation.id);
                await adapter.continueConversation(reference, async (context) => {
                    await context.sendActivity(req.body);
                });
            } catch (err) {
                console.error(`Error sending to conversation ${reference.conversation.id}:`, err);
            }
        }

        res.status(200).send('Notifications sent successfully');
    } catch (error) {
        console.error('Webhook error:', error);
        res.status(500).send(error.message);
    }
});

const port = process.env.PORT || 3978;
app.listen(port, () => {
    console.log(`\n${bot.constructor.name} listening at http://localhost:${port}`);
});
