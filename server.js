import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";
import multer from "multer";
import fs from "fs-extra";

dotenv.config();

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// -------- FILE UPLOADS ----------
const upload = multer({ dest: "uploads/" });

// -------- MSAL CONFIG ----------
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET
  },
};

const REDIRECT_URI = process.env.REDIRECT_URI;
const SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.Send",
  "offline_access"
];

const pca = new msal.ConfidentialClientApplication(msalConfig);

// Store tokens
let userTokens = {};

app.get("/", (req, res) => {
  res.send("📧 APEX Mail API is running.");
});

// ---------- LOGIN ----------
app.get("/login", async (req, res) => {
  const userEmail = req.query.userEmail;
  if (!userEmail) return res.status(400).send("Missing userEmail");

  const authUrl = await pca.getAuthCodeUrl({
    scopes: SCOPES,
    redirectUri: REDIRECT_URI,
    state: encodeURIComponent(userEmail)
  });

  res.redirect(authUrl);
});

// ---------- REDIRECT ----------
app.get("/redirect", async (req, res) => {
  const code = req.query.code;
  const userEmail = decodeURIComponent(req.query.state);

  const tokenResponse = await pca.acquireTokenByCode({
    code,
    scopes: SCOPES,
    redirectUri: REDIRECT_URI
  });

  userTokens[userEmail] = {
    accessToken: tokenResponse.accessToken
  };

  res.send(`✅ ${userEmail} authenticated! Ready to send mails.`);
});

// ---------- REFRESH TOKEN ----------
async function getAccessToken(userEmail) {
  try {
    const accounts = await pca.getTokenCache().getAllAccounts();
    const account = accounts.find(a => a.username === userEmail);
    if (!account) return null;

    const res = await pca.acquireTokenSilent({
      account,
      scopes: SCOPES
    });

    return res.accessToken;
  } catch (err) {
    console.log("Refresh error:", err.message);
    return null;
  }
}

// ---------- SEND MAIL (CORE FEATURE) ----------
app.post(
  "/send-mail",
  upload.array("attachments"),
  async (req, res) => {
    const {
      sender_email,
      to_email,
      cc_email,
      bcc_email,
      subject,
      body_html,
      template_id,
      signature_html,
      lead_id
    } = req.body;

    let accessToken = await getAccessToken(sender_email);
    if (!accessToken)
      return res.status(401).json({
        error: `User ${sender_email} not authenticated. Login first.`
      });

    // ---------- LOAD TEMPLATE IF ANY ----------
    let finalBody = body_html || "";
    if (template_id) {
      const templates = {
        1: "<h2>Welcome!</h2><p>This is Template 1.</p>",
        2: "<h2>Offer Mail</h2><p>This is Template 2.</p>"
      };
      finalBody = templates[template_id] + finalBody;
    }

    // ---------- ADD SIGNATURE ----------
    if (signature_html) {
      finalBody += `<br><br>${signature_html}`;
    }

    // ---------- ATTACHMENTS ----------
    let attachments = [];

    for (let file of req.files) {
      const content = await fs.readFile(file.path, { encoding: "base64" });

      attachments.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: file.originalname,
        contentBytes: content
      });

      await fs.remove(file.path);
    }

    // ---------- GRAPH MAIL BODY ----------
    const mail = {
      message: {
        subject,
        body: { contentType: "HTML", content: finalBody },
        toRecipients: [{ emailAddress: { address: to_email } }],
        ccRecipients: cc_email ? [{ emailAddress: { address: cc_email } }] : [],
        bccRecipients: bcc_email ? [{ emailAddress: { address: bcc_email } }] : [],
        attachments
      },
      saveToSentItems: "true"
    };

    // ---------- SEND MAIL ----------
    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(mail)
    });

    const text = await response.text();
    if (!response.ok) return res.status(400).json({ error: text });

    // ---------- SAVE HISTORY ----------
    console.log(`📨 EMAIL SENT (lead_id=${lead_id})`);

    res.json({ success: true, message: "Mail sent successfully", lead_id });
  }
);

// ---------- START SERVER ----------
const PORT = process.env.PORT || 7000;
app.listen(PORT, () =>
  console.log(`🚀 APEX Mail API running on port ${PORT}`)
);
