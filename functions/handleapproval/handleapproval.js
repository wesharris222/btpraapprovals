const fetch = require('node-fetch');

module.exports = async function (context, req) {
    try {
        context.log('handleapproval function started');
        context.log('Request query:', JSON.stringify(req.query));

        const decision = req.query.decision;
        const message = req.query.message || 'default';
        let approvalUrl = req.query.approvalUrl;
        const authKey = req.query.authKey;

        if (!approvalUrl || !authKey) {
            throw new Error('Missing required parameters: approvalUrl or authKey');
        }

        // Ensure proper URL construction
        if (!approvalUrl.startsWith('http')) {
            approvalUrl = `https://${approvalUrl}`;
        }
        
        // Add /approve_jump_request if not present
        if (!approvalUrl.endsWith('/approve_jump_request')) {
            approvalUrl = `${approvalUrl}/approve_jump_request`;
        }

        // Create the proper URL with authKey parameter
        const fullUrl = `${approvalUrl}?authKey=${authKey}`;

        // Create form data for the request
        const formData = new URLSearchParams();
        formData.append('authKey', authKey);
        formData.append('comments', message);
        
        // Add the appropriate button value based on decision
        if (decision === 'approved') {
            formData.append('approved', 'Approve');
        } else {
            formData.append('denied', 'Deny');
        }

        context.log('Sending request to:', fullUrl);
        context.log('Form data:', formData.toString());

        // Send the approval/denial request
        const response = await fetch(fullUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            },
            body: formData.toString()
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Approval request failed with status: ${response.status}. Response: ${errorText}`);
        }

        const responseText = await response.text();
        context.log('Approval response:', responseText);

        context.res = {
            status: 200,
            headers: {
                'Content-Type': 'application/json'
            },
            body: {
                message: `Request ${decision} successfully processed`,
                details: responseText
            }
        };
        
    } catch (error) {
        context.log.error('Error in handleapproval function:', error);
        context.res = {
            status: 500,
            headers: {
                'Content-Type': 'application/json'
            },
            body: {
                error: error.message
            }
        };
    }
};