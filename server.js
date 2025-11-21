// server.js
import express from "express";
import * as msal from "@azure/msal-node";
import fetch from "node-fetch";
import dotenv from "dotenv";
import multer from "multer";
import fs from "fs-extra";
import path from "path";

dotenv.config();

const app = express();
app.use(express.json({ limit: "12mb" }));
app.use(express.urlencoded({ extended: true }));

const upload = multer({ dest: "uploads/" });

// ---------- MSAL config ----------
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const REDIRECT_URI = process.env.REDIRECT_URI;
const SCOPES = [
  "https://graph.microsoft.com/User.Read",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Calendars.ReadWrite",
  "offline_access",
];

const pca = new msal.ConfidentialClientApplication(msalConfig);

// ---------- token persistence file ----------
const TOKENS_FILE = path.join(process.cwd(), "tokens.json");
async function loadTokens() {
  try {
    const exists = await fs.pathExists(TOKENS_FILE);
    if (!exists) return {};
    const txt = await fs.readFile(TOKENS_FILE, "utf8");
    return JSON.parse(txt || "{}");
  } catch (e) {
    console.error("loadTokens error", e);
    return {};
  }
}
async function saveTokens(tokens) {
  try {
    await fs.writeFile(TOKENS_FILE, JSON.stringify(tokens, null, 2), "utf8");
  } catch (e) {
    console.error("saveTokens error", e);
  }
}

// in-memory cache mirrored with tokens.json
let userTokens = {};
(async () => { userTokens = await loadTokens(); })();

// ---------- helpers ----------
function nowTs() { return Math.floor(Date.now() / 1000); } // seconds

async function getAccessToken(userEmail) {
  const email = (userEmail || "").toLowerCase();
  const tokens = userTokens[email];
  if (!tokens) {
    console.log("No tokens stored for", email);
    return null;
  }
  // check expiry (expiresOn stored as seconds)
  if (tokens.expiresOn && tokens.expiresOn > nowTs() + 30) {
    return tokens.accessToken;
  }
  // expired — cannot reliably refresh in demo; ask user to re-login
  console.log("Access token expired for", email);
  return null;
}

// ---------- root ----------
app.get("/", (req, res) => {
  res.send("APEX Mail API running. Use /login?userEmail=you@domain.com");
});

// ---------- login (start auth) ----------
app.get("/login", async (req, res) => {
  const userEmail = req.query.userEmail;
  if (!userEmail) return res.status(400).send("Missing userEmail");
  try {
    const authUrl = await pca.getAuthCodeUrl({
      scopes: SCOPES,
      redirectUri: REDIRECT_URI,
      state: encodeURIComponent(userEmail),
    });
    res.redirect(authUrl);
  } catch (err) {
    console.error("getAuthCodeUrl error", err);
    res.status(500).send("Error creating auth URL");
  }
});

// ---------- redirect (finish auth) ----------
app.get("/redirect", async (req, res) => {
  const code = req.query.code;
  const userEmail = decodeURIComponent(req.query.state || "");
  if (!code || !userEmail) return res.status(400).send("Missing code or userEmail");

  try {
    const tokenRequest = {
      code,
      scopes: SCOPES,
      redirectUri: REDIRECT_URI,
    };

    const result = await pca.acquireTokenByCode(tokenRequest);

    // store tokens in file + memory (simple demo persistence)
    const email = (userEmail || result.account?.username || "").toLowerCase();
    userTokens[email] = {
      accessToken: result.accessToken,
      refreshToken: result.refreshToken || null,
      expiresOn: result.expiresOn ? Math.floor(new Date(result.expiresOn).getTime() / 1000) : nowTs() + 3600,
      account: result.account || null
    };
    await saveTokens(userTokens);

    console.log("Token saved for", email);
    res.send(`<h2>Authenticated: ${email}</h2><p>You can now call /send-mail</p>`);

  } catch (err) {
    console.error("redirect error", err);
    res.status(500).send("Error acquiring token: " + (err.message || err));
  }
});

// ---------- debug accounts ----------
app.get("/debug-accounts", async (req, res) => {
  const tokens = await loadTokens();
  res.json({ stored_accounts: Object.keys(tokens) });
});

// ---------- send-mail (supports JSON attachments base64 OR multipart files) ----------
app.post("/send-mail", upload.array("attachments"), async (req, res) => {
  try {
    // support both JSON and form-data
    const isMultipart = req.files && req.files.length > 0;
    let payload = req.body;

    // if content-type application/json, body already parsed
    // expected JSON keys: sender_email, toEmails (array), ccEmails (array), bccEmails (array),
    // subject, body (html), template_html, signature_html, attachments [{ filename, contentBase64 }]
    if (!isMultipart && req.headers["content-type"] && req.headers["content-type"].includes("application/json")) {
      payload = req.body;
    }

    const sender_email = (payload.sender_email || "").toLowerCase();
    if (!sender_email) return res.status(400).json({ error: "sender_email required" });

    const accessToken = await getAccessToken(sender_email);
    if (!accessToken) {
      return res.status(401).json({ error: `User ${sender_email} not authenticated. Visit /login?userEmail=${sender_email}` });
    }

    const toEmails = payload.toEmails || (payload.toEmailsCSV ? payload.toEmailsCSV.split(",").map(s=>s.trim()) : []);
    const ccEmails = payload.ccEmails || (payload.ccEmailsCSV ? payload.ccEmailsCSV.split(",").map(s=>s.trim()) : []);
    const bccEmails = payload.bccEmails || (payload.bccEmailsCSV ? payload.bccEmailsCSV.split(",").map(s=>s.trim()) : []);
    const subject = payload.subject || "(no subject)";
    const bodyHtml = payload.body || payload.body_html || "";
    const templateHtml = payload.template_html || "";
    const signatureHtml = payload.signature_html || "";

    let finalBody = templateHtml + "<br/>" + bodyHtml + "<br/><br/>" + signatureHtml;

    // attachments: either multipart files (req.files) OR payload.attachments (JSON array with base64)
    let attachments = [];

    if (isMultipart) {
      for (const f of req.files) {
        const content = await fs.readFile(f.path, { encoding: "base64" });
        attachments.push({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: f.originalname,
          contentBytes: content
        });
        // remove temp file
        await fs.remove(f.path);
      }
    } else if (payload.attachments && Array.isArray(payload.attachments)) {
      for (const a of payload.attachments) {
        if (a.filename && a.contentBase64) {
          attachments.push({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: a.filename,
            contentBytes: a.contentBase64
          });
        }
      }
    }

    // build graph payload
    const mail = {
      message: {
        subject,
        body: { contentType: "HTML", content: finalBody },
        toRecipients: (toEmails || []).map(address => ({ emailAddress: { address } })),
        ccRecipients: (ccEmails || []).map(address => ({ emailAddress: { address } })),
        bccRecipients: (bccEmails || []).map(address => ({ emailAddress: { address } })),
        attachments
      },
      saveToSentItems: true
    };

    const graphResp = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(mail)
    });

    const text = await graphResp.text();
    if (!graphResp.ok) {
      console.error("graph error", text);
      return res.status(400).json({ error: "sendMail failed", details: text });
    }

    // Optionally persist history (not included here) — APEX will save to DB
    res.json({ success: true, message: "Email sent" });
  } catch (err) {
    console.error("send-mail error", err);
    res.status(500).json({ error: err.message || err });
  }
});

// ---------- create meeting (same as earlier) ----------
app.post("/create-meeting", async (req, res) => {
  try {
    const sender_email = (req.body.sender_email || "").toLowerCase();
    if (!sender_email) return res.status(400).json({ error: "Missing sender_email" });

    const accessToken = await getAccessToken(sender_email);
    if (!accessToken) return res.status(401).json({ error: `User ${sender_email} not authenticated. Visit /login?userEmail=${sender_email}` });

    const attendees = (req.body.attendees || []).map(email => ({ emailAddress: { address: email, name: email }, type: "required" }));

    const event = {
      subject: req.body.subject || "Meeting from Oracle APEX",
      body: { contentType: "HTML", content: req.body.description || "Meeting via Oracle APEX" },
      start: { dateTime: req.body.start, timeZone: req.body.timeZone || "India Standard Time" },
      end: { dateTime: req.body.end, timeZone: req.body.timeZone || "India Standard Time" },
      location: { displayName: req.body.location || "Online" },
      attendees,
      isOnlineMeeting: true,
      onlineMeetingProvider: "teamsForBusiness",
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/events", {
      method: "POST",
      headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
      body: JSON.stringify(event)
    });

    const result = await response.json();
    if (!response.ok) return res.status(400).json({ error: "Failed to create meeting", details: result });

    res.json({ success: true, eventId: result.id, joinUrl: result.onlineMeeting?.joinUrl || null });
  } catch (err) {
    console.error("create-meeting error", err);
    res.status(500).json({ error: err.message || err });
  }
});

// ---------- start ----------
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`APEX Mail API listening on ${PORT}`));
