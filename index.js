import { ConfidentialClientApplication } from '@azure/msal-node'
import { Client } from '@microsoft/microsoft-graph-client'
import express from 'express'
import session from 'express-session'
import dotenv from 'dotenv'
dotenv.config()

const HOST = process.env.HOST || 'http://localhost:3000'

const msalClient = new ConfidentialClientApplication({
    auth: {
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        authority: 'https://login.microsoftonline.com/common'
    }
})

function getAuthenticatedClient(msalClient, userId) {
    const client = Client.init({
        // Implement an auth provider that gets a token
        // from the app's MSAL instance
        authProvider: async (done) => {
            try {
                // Get the user's account
                const account = await msalClient
                    .getTokenCache()
                    .getAccountByHomeId(userId);

                if (account) {
                    const response = await msalClient.acquireTokenSilent({
                        scopes: ['user.read', 'mail.read'],
                        redirectUri: `${HOST}/auth/callback`,
                        account: account
                    });

                    done(null, response.accessToken);
                }
            } catch (err) {
                console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
                done(err, null);
            }
        }
    });

    return client;
}

const graph = {
    getUserDetails: async function (msalClient, userId) {
        const client = getAuthenticatedClient(msalClient, userId);

        const user = await client
            .api('/me')
            .select('displayName,userPrincipalName')
            .get();
        return user;
    },

    getEmails: async function (msalClient, userId) {
        const client = getAuthenticatedClient(msalClient, userId);

        const messages = await client
            .api('/me/messages')
            .select('sender,subject,from,toRecipients')
            .get();
        return messages;
    },

    getEmailById: async function (msalClient, userId, emailId) {
        const client = getAuthenticatedClient(msalClient, userId);

        const message = await client
            .api(`/me/messages/${emailId}`)
            .select('sender,subject,from,toRecipients')
            .get();
        return message;
    },

    createSubcription: async function (msalClient, userId) {
        const client = getAuthenticatedClient(msalClient, userId);

        await client.api('/subscriptions')
            .post({
                changeType: 'created,updated',
                notificationUrl: `${HOST}/auth/notification`,
                lifecycleNotificationUrl: `${HOST}/auth/notification`,
                resource: '/me/messages',
                expirationDateTime: '2024-05-20',
                clientState: process.env.CLIENT_STATE
            })
    },

    getSubcription: async function (msalClient, userId) {
        const client = getAuthenticatedClient(msalClient, userId);
        const data = await client.api('/subscriptions').get()
        return data.value
    },

    deleteAllSubcription: async function (msalClient, userId) {
        const existingSubcriptions = await this.getSubcription(msalClient, userId)
        await Promise.allSettled(existingSubcriptions.map(async sub => {
            const client = getAuthenticatedClient(msalClient, userId);
            const id = sub.id
            return await client.api(`/subscriptions/${id}`).delete()
        }))
    }
}

const app = express()
app.use(session({
    secret: 'super secret',
    resave: false,
    saveUninitialized: false,
    unset: 'destroy'
}))
app.use(express.json())

app.get('/auth/signin', async (req, res) => {
    const urlParameters = {
        scopes: ['user.read', 'mail.read'],
        redirectUri: `${HOST}/auth/callback`
    }

    const authUrl = await msalClient.getAuthCodeUrl(urlParameters)
    res.json({ redirect: authUrl })
})

app.get('/auth/callback', async (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ['user.read', 'mail.read'],
        redirectUri: `${HOST}/auth/callback`
    };

    const response = await msalClient.acquireTokenByCode(tokenRequest)

    // save response as a session
    req.session.userId = response.account.homeAccountId

    await graph.deleteAllSubcription(msalClient, req.session.userId)

    await graph.createSubcription(
        msalClient,
        req.session.userId
    );
    console.log(`Create a subcription successfully.`)

    res.json(response)
})

app.post('/auth/notification', async (req, res) => {
    if (req.query.validationToken) {
        const validationToken = req.query.validationToken
        res.status(200).type('text/plain').send(validationToken)
    }
    else {
        const changedData = req.body.value

        // log new created/updated emails
        const newEmails = await Promise.all(changedData.map(async data => {
            const convertedData = JSON.parse(JSON.stringify(data))
            console.log(convertedData)
            if (convertedData.clientState === process.env.CLIENT_STATE &&
                convertedData.resourceData['@odata.type'] === "#Microsoft.Graph.Message") {
                return await graph.getEmailById(
                    msalClient,
                    req.session.userId,
                    convertedData.resourceData.id)
            }
            return
        }))
        console.log(newEmails)
        res.status(200).send('OK')
    }
})

app.get('/user', async (req, res) => {
    if (!req.session.userId) {
        res.status(400).send('Error')
    }
    const user = await graph.getUserDetails(
        msalClient,
        req.session.userId
    );

    res.json(user)
})

app.get('/emails', async (req, res) => {
    if (!req.session.userId) {
        res.status(400).send('Error')
    }
    const emails = await graph.getEmails(
        msalClient,
        req.session.userId
    );

    res.json(emails)
})

app.get('/emails/:id', async (req, res) => {
    if (!req.session.userId) {
        res.status(400).send('Error')
    }
    const email = await graph.getEmailById(
        msalClient,
        req.session.userId,
        req.params.id
    );

    res.json(email)
})

app.listen(3000, () => {
    console.log('express starts on port 3000')
})
