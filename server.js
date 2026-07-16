// Vapi -> Google Sheets webhook server
// Receives the "end-of-call-report" from Vapi and writes a clean row to Google Sheets.

const express = require("express");
const { google } = require("googleapis");

const app = express();
app.use(express.json({ limit: "5mb" }));

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1";

// Google auth using a service account (set up instructions in README.md)
// Preferred: GOOGLE_CREDENTIALS_BASE64 - the whole service account JSON file, base64 encoded.
// This avoids all copy-paste issues with newlines/quotes in the private key.
// Fallback: separate GOOGLE_CLIENT_EMAIL + GOOGLE_PRIVATE_KEY variables.
let credentials;
if (process.env.GOOGLE_CREDENTIALS_BASE64) {
  const decoded = Buffer.from(process.env.GOOGLE_CREDENTIALS_BASE64, "base64").toString("utf8");
  const parsed = JSON.parse(decoded);
  credentials = {
    client_email: parsed.client_email,
    private_key: parsed.private_key,
  };
} else {
  credentials = {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
  };
}

const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

// Simple health check so you can confirm the server is alive in a browser
app.get("/", (req, res) => {
  res.send("Vapi webhook server is running.");
});

app.post("/webhook", async (req, res) => {
  try {
    const body = req.body;
    const message = body.message || body; // Vapi nests everything under "message"

    // Only act on the final call report - ignore status updates, transcripts, etc.
    if (message.type !== "end-of-call-report") {
      return res.status(200).send("Ignored (not end-of-call-report)");
    }

    const customer = message.customer || {};
    const analysis = message.analysis || {};
    const call = message.call || {};

    // Try a few likely places Vapi puts the caller's name (varies by setup)
    const name =
      analysis.structuredData?.name ||
      customer.name ||
      "Unknown";

    const phone = analysis.structuredData?.phone || customer.number || "Unknown";

    const email =
      analysis.structuredData?.email ||
      "Unknown";

    const reason = analysis.structuredData?.reason || "Unknown";

    const summary = message.summary || analysis.summary || "No summary";

    // Row order matches sheet headers: name | phone | email | reason for call | call summary
    const row = [name, phone, email, reason, summary];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A:E`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: { values: [row] },
    });

    console.log("Row added:", row);
    res.status(200).send("OK");
  } catch (err) {
    console.error("Error handling webhook:", err);
    res.status(500).send("Server error");
  }
});

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});

