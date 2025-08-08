const express = require('express');
const session = require('express-session');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const cors = require('cors');
const { GoogleGenerativeAI } = require('@google/generative-ai');
require('isomorphic-fetch');
require('dotenv').config();

const app = express();
const PORT = 3000;

// --- Middleware & Config ---
app.use(express.json()); 
app.use(cors({
    origin: ['http://localhost:3001', 'https://siteweave-ai.netlify.app'], // Allow both local and live frontend
    credentials: true
}));

app.use(session({
    secret: process.env.SESSION_SECRET || 'a-super-secret-key-for-your-project-brain',
    resave: false,
    saveUninitialized: false,
    cookie: { secure: process.env.NODE_ENV === 'production', httpOnly: true, sameSite: 'none' }
}));

// --- DYNAMIC REDIRECT URI ---
const isProduction = process.env.NODE_ENV === 'production';
const redirectUri = isProduction 
    ? "https://siteweave-ai-backend.onrender.com/redirect" 
    : "http://localhost:3000/redirect";

console.log(`Redirect URI set to: ${redirectUri}`); // For debugging

const msalConfig = {
    auth: {
        clientId: process.env.MS_CLIENT_ID,
        authority: `https://login.microsoftonline.com/common`,
        clientSecret: process.env.MS_CLIENT_SECRET,
    }
};
const pca = new msal.ConfidentialClientApplication(msalConfig);
const scopes = ["Mail.Read", "User.Read", "offline_access"];


// --- Gemini API Integration ---
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: "gemini-pro" });

async function getGeminiAnalysis(emailContent) {
    console.log("Sending content to Gemini for real analysis...");
    const prompt = `
        Analyze the following construction-related email. Extract key information and respond ONLY with a valid JSON object.
        The JSON object must have these exact keys: "summary", "actionItems", "sentiment", and "docType".
        - "summary": A concise, professional summary of the email's main purpose.
        - "actionItems": An array of strings, with each string being a specific, actionable task, question, or deadline. If no action items, return an empty array.
        - "sentiment": Classify the sentiment as "Positive", "Negative", or "Neutral".
        - "docType": Classify the document type (e.g., "RFI", "Change Order", "Submittal", "Invoice", "General Correspondence").

        Email Content:
        ---
        ${emailContent}
        ---
    `;
    try {
        const result = await model.generateContent(prompt);
        const response = await result.response;
        const text = response.text();
        const jsonString = text.replace(/```json/g, '').replace(/```/g, '').trim();
        return JSON.parse(jsonString);
    } catch (error) {
        console.error("Error calling Gemini API:", error);
        return {
            summary: "AI analysis failed for this email.",
            actionItems: [],
            sentiment: "Neutral",
            docType: "Unknown"
        };
    }
}

// --- Authentication & API Routes ---

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

app.get('/me', (req, res) => {
    if (req.session.account) {
        res.status(200).json({ loggedIn: true, account: req.session.account });
    } else {
        res.status(200).json({ loggedIn: false });
    }
});

app.post('/fetch-emails', async (req, res) => {
    if (!req.session.account) {
        return res.status(401).send("User not authenticated.");
    }
    const { searchQuery } = req.body;
    try {
        const tokenRequest = { account: req.session.account, scopes: scopes };
        const tokenResponse = await pca.acquireTokenSilent(tokenRequest);
        const accessToken = tokenResponse.accessToken;
        const graphClient = Client.init({ authProvider: (done) => done(null, accessToken) });
        let messagesRequest = graphClient.api('/me/messages').select('id,subject,from,receivedDateTime,body').top(10);
        if (searchQuery) {
            messagesRequest = messagesRequest.search(`"${searchQuery}"`);
        }
        const messages = await messagesRequest.get();
        const analysisPromises = messages.value.map(async (email) => {
            const cleanBody = email.body.content.replace(/<[^>]*>?/gm, ''); 
            const contentToAnalyze = `Subject: ${email.subject}\nBody: ${cleanBody}`;
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