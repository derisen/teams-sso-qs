const express = require('express');
const msal = require('@azure/msal-node');

const app = express();

require('dotenv').config();

// Before running the sample, you will need to replace the values in the config, 
// including the clientSecret
const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        clientSecret:  process.env.CLIENT_SECRET,
    }
};

// Create msal application object
const cca = new msal.ConfidentialClientApplication(config);

app.get('/getGraphAccessToken', async (req,res) => {

    // TODO: access token validation

    const oboRequest = {
        oboAssertion: req.query.ssoToken,
        scopes: [process.env.GRAPH_SCOPES],
    }

    try {
        let response = await cca.acquireTokenOnBehalfOf(oboRequest);
        console.log(response);
        res.send(response);   
    } catch (error) {
        console.log(error)
        res.send(error);
    }

    // if(!response.ok) {
    //     if( (data.error === 'invalid_grant') || (data.error === 'interaction_required') ) {
    //         //This is expected if it's the user's first time running the app ( user must consent ) or the admin requires MFA
    //         console.log("User must consent or perform MFA. You may also encouter this error if your client ID or secret is incorrect.");
    //         res.status(403).json({ error: 'consent_required' }); //This error triggers the consent flow in the client.
    //     } else {
    //         //Unknown error
    //         console.log('Could not exchange access token for unknown reasons.');
    //         res.status(500).json({ error: 'Could not exchange access token' });
    //     }
    // } else {
    //     //The on behalf of token exchange worked. Return the token (data object) to the client.
    //     res.send(data);
    // }
});

const port = process.env.PORT || 5000;

app.listen(port);

console.log('API server is listening on port ' + port);
