const express = require('express');
const session = require('express-session');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const cors = require('cors');
require('isomorphic-fetch');
require('dotenv').config();

const app = express();
const PORT = 3000;

// --- Middleware Setup ---
app.use(cors({
    origin: 'http://localhost:3001',
    credentials: true
}));

app.use(session({
    secret: process.env.SESSION_SECRET || 'a-super-secret-key-for-your-project-brain',
    resave: false,
    saveUninitialized: false,
    cookie: {
        secure: false, 
    }
}));

// --- MSAL Configuration ---
const msalConfig = {
    auth: {
        clientId: process.env.MS_CLIENT_ID,
        // Using 'common' endpoint to allow personal and work/school accounts
        authority: `https://login.microsoftonline.com/common`, 
        clientSecret: process.env.MS_CLIENT_SECRET,
    }
};

const pca = new msal.ConfidentialClientApplication(msalConfig);
const scopes = ["Mail.Read", "User.Read", "offline_access"]; 
const redirectUri = "http://localhost:3000/redirect";

// --- Gemini API Placeholder ---
async function getGeminiAnalysis(emailContent) {
    console.log("Sending content to Gemini for analysis...");
    await new Promise(resolve => setTimeout(resolve, 500));
    return {
        summary: "This is a critical RFI regarding a structural beam discrepancy on drawing S-102.",
        actionItems: ["Consult structural engineer.", "Provide cost impact by EOD."],
        sentiment: "Negative",
        docType: "RFI"
    };
}

// --- Authentication Routes ---
app.get('/login', (req, res) => {
    const authCodeUrlParameters = { scopes, redirectUri };
    pca.getAuthCodeUrl(authCodeUrlParameters)
        .then((response) => res.redirect(response))
        .catch((error) => res.status(500).send(error));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = { code: req.query.code, scopes, redirectUri };
    pca.acquireTokenByCode(tokenRequest)
        .then((response) => {
            req.session.account = response.account; 
            res.send('<script>window.close();</script>');
        })
        .catch((error) => res.status(500).send(error));
});

app.get('/logout', (req, res) => {
    req.session.destroy(() => {
        res.status(200).send('Successfully logged out');
    });
});

// --- API Routes ---
app.get('/me', (req, res) => {
    if (req.session.account) {
        res.status(200).json({ loggedIn: true, account: req.session.account });
    } else {
        res.status(200).json({ loggedIn: false });
    }
});

app.get('/fetch-emails', async (req, res) => {
    if (!req.session.account) {
        return res.status(401).send("User not authenticated.");
    }

    try {
        const tokenRequest = {
            account: req.session.account,
            scopes: scopes,
        };
        
        const tokenResponse = await pca.acquireTokenSilent(tokenRequest);
        const accessToken = tokenResponse.accessToken;

        const graphClient = Client.init({
            authProvider: (done) => done(null, accessToken)
        });

        const messages = await graphClient
            .api('/me/messages')
            .select('id,subject,from,receivedDateTime,bodyPreview')
            .top(5)
            .get();

        const analysisPromises = messages.value.map(async (email) => {
            const contentToAnalyze = `Subject: ${email.subject}\nBody: ${email.bodyPreview}`;
            const analysis = await getGeminiAnalysis(contentToAnalyze);
            return { ...email, analysis };
        });

        const analyzedEmails = await Promise.all(analysisPromises);
        res.status(200).json(analyzedEmails);

    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            return res.status(401).send("Session expired. Please log in again.");
        }
        console.error(error);
        res.status(500).send("Error fetching or analyzing emails.");
    }
});

app.listen(PORT, () => console.log(`Backend server running on http://localhost:${PORT}`));
