const { TeamsActivityHandler, MessageFactory } = require('botbuilder');
const { TableClient } = require('@azure/data-tables');
const fetch = require('node-fetch');

class BTApprovalBot extends TeamsActivityHandler {
    constructor() {
        super();
        
        this.tableClient = null;
        this.initializeStorage().then(() => {
            console.log('Storage initialized successfully');
        }).catch(err => {
            console.error('Error initializing storage:', err);
        });
    }

    async initializeStorage() {
        try {
            this.tableClient = TableClient.fromConnectionString(
                process.env.AzureStorageConnectionString,
                'conversationreferences'
            );
            await this.tableClient.createTable();
            console.log('Storage table created or exists');
        } catch (err) {
            console.error('Error creating table:', err);
            throw err;
        }
    }

    async onInstallationUpdate(context) {
        console.log('Installation update activity:', context.activity);
        if (context.activity.action === 'add') {
            await this.addConversationReference(context.activity);
            await context.sendActivity("Hi! I'm the BeyondTrust PRA approvals bot. I'll notify you of any approval requests.");
        }
    }

    async onConversationUpdateActivity(context) {
        await this.addConversationReference(context.activity);
        
        if (context.activity.membersAdded && context.activity.membersAdded.length > 0) {
            for (let idx in context.activity.membersAdded) {
                if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                    await context.sendActivity("Hi! I'm the BeyondTrust PRA approvals bot. I'll notify you of any approval requests.");
                }
            }
        }
        
        await super.onConversationUpdateActivity(context);
    }

    async onInvokeActivity(context) {
        console.log('Invoke Activity:', context.activity);

        if (context.activity.name === 'adaptiveCard/action') {
            const actionData = context.activity.value.action.data;
            console.log('Action Data:', actionData);
            
            try {
                const functionUrl = process.env.FUNCTIONAPP_URL;
                const functionKey = process.env.FUNCTIONAPP_KEY;

                const message = actionData.approval_message || "Not specified";
                
                let duration = "Once";
                if (actionData.duration_type === "seconds" && actionData.duration_seconds) {
                    duration = actionData.duration_seconds.toString();
                }

                const username = context.activity.from.name || 'Unknown User';
                
                const functionParams = new URLSearchParams({
                    decision: actionData.decision,
                    requestId: actionData.requestId,
                    ticketId: actionData.ticketNumber,
                    message: message,
                    duration: duration,
                    username: username,
                    approvalUrl: actionData.approvalUrl,
                    authKey: actionData.authKey
                }).toString();

                console.log('Calling function with params:', functionParams);

                const response = await fetch(`${functionUrl}?${functionParams}`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'x-functions-key': functionKey
                    }
                });

                const responseData = await response.json();
                console.log('Function Response:', responseData);

                if (response.ok) {
                    return {
                        status: 200,
                        body: {
                            statusCode: 200,
                            type: 'application/vnd.microsoft.activity.message',
                            value: `Request ${actionData.decision} successfully processed by ${username}.`
                        }
                    };
                } else {
                    throw new Error(`Function call failed: ${JSON.stringify(responseData)}`);
                }
            } catch (error) {
                console.error('Error processing action:', error);
                return {
                    status: 500,
                    body: {
                        statusCode: 500,
                        type: 'application/vnd.microsoft.activity.message',
                        value: `Error: ${error.message}`
                    }
                };
            }
        }
        return null;
    }

    async addConversationReference(activity) {
        if (!activity?.conversation?.id) {
            console.log('Invalid activity format:', activity);
            return;
        }

        try {
            if (!this.tableClient) {
                await this.initializeStorage();
            }

            const entity = {
                partitionKey: 'channel',
                rowKey: activity.conversation.id,
                reference: JSON.stringify({
                    channelId: activity.channelId,
                    serviceUrl: activity.serviceUrl,
                    conversation: {
                        id: activity.conversation.id,
                        name: activity.conversation.name,
                        conversationType: activity.conversation.conversationType,
                        isGroup: activity.conversation.isGroup,
                        tenantId: activity.conversation.tenantId
                    },
                    bot: activity.recipient,
                    tenantId: activity.conversation.tenantId
                })
            };

            await this.tableClient.upsertEntity(entity);
            console.log('Stored conversation reference');
        } catch (err) {
            console.error('Error storing conversation reference:', err);
        }
    }

    async getAllConversationReferences() {
        try {
            if (!this.tableClient) {
                await this.initializeStorage();
            }

            const references = [];
            const entities = this.tableClient.listEntities({
                queryOptions: { filter: "PartitionKey eq 'channel'" }
            });

            for await (const entity of entities) {
                if (entity.reference) {
                    try {
                        const reference = JSON.parse(entity.reference);
                        references.push(reference);
                    } catch (err) {
                        console.error('Error parsing reference:', err);
                    }
                }
            }

            return references;
        } catch (err) {
            console.error('Error retrieving conversation references:', err);
            return [];
        }
    }
}

module.exports = BTApprovalBot;
